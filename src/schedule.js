/**
 * @param {(idleDeadline: IdleDeadline) => void} callback
 */
export const schedule = globalThis.requestIdleCallback
    || (() => {
        // This fallback uses a `MessageChannel` and `postMessage()` technique
        // to trigger a task asynchronously without the delays or interferences
        // that might occur with `setTimeout()` or `setImmediate()`
        const fps = 1e3 / 60 // time which can be used without event loop lags
        return function (callback) {
            const channel = new MessageChannel() // create new channel at each time is required to fix behavior in Safari
            // prefer `addEventListener()` because `onmessage` is not for handling once
            channel.port1.addEventListener('message', () => { // handler runs asynchronously when message is come
                // note that each call of this handler is queued by execution environment because using immediate posting of message
                // so all callbacks runs strict step-by-step and there is no need to be careful about interferences in execution
                const frameDeadline = performance.now() + fps // compute deadline
                callback({ // execute next task
                    didTimeout: false, // is not supported yet, so just pass a stub...
                    timeRemaining: () => frameDeadline - performance.now() // returns remaining time
                })
            }, {
                once: true // allow to collect channel by GC after task execution
                           // it is required in some environments like Node.js
            })
            channel.port2.postMessage(0) // post a message to trigger a callback asynchronously
        }
    })()
    // || (() => callback => {
    //     const fps = 1e3 / 60 // time which can be used without event loop lags
    //     const frameDeadline = performance.now() + fps
    //     callback({
    //         didTimeout: false, // is not supported yet, so just pass a stub...
    //         timeRemaining: () => frameDeadline - performance.now() // returns remaining time
    //     })
    // })()
