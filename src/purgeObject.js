import { isArray } from "./isArray.js"
import { isObject } from "./isObject.js"

/**
 * @param {{ [key: string]: unknown }} a 
 * @returns {{ [key: string]: unknown }}
 */
export function purgeObject(a) {
    return Object.fromEntries(
        Object.entries(a)
            .filter(entry => entry[1] !== undefined)
            .map(([ key, value ]) => [
                key,
                isObject(value) && !isArray(value) && value !== null
                    ? purgeObject(/** @type {{ [key: string]: unknown }} */ (value))
                    : value
            ])
    )
}
