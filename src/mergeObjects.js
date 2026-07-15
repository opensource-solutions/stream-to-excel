import { isArray } from "./isArray.js"
import { isObject } from "./isObject.js"

/**
 * 
 * @param {{ [key: string]: unknown }} a 
 * @param {{ [key: string]: unknown }} b 
 * @returns {{ [key: string]: unknown }}
 */
export function mergeObjects(a, b) {
    const result = { ...a, ...b }
    Object.entries(b).forEach(({ 0: k, 1: v }) => {
        if (isObject(v) && !isArray(v) && v !== null) {
            let c = a[k]
            if (!isObject(c) || isArray(c) || c === null)
                c = {}
            result[k] = mergeObjects(
                /** @type {{ [key: string]: unknown }} */ (c),
                /** @type {{ [key: string]: unknown }} */ (v)
            )
        }
    })
    return result
}
