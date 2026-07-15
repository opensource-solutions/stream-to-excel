import { mergeObjects } from './mergeObjects.js'
import { purgeObject } from './purgeObject.js'

/**
 * @typedef {import('../types/index.d.ts').CellStyle} CellStyle
 * @typedef {import('../types/index.d.ts').CellStyleOptions} CellStyleOptions
 */

/**
 * @param {CellStyle} cellStyle
 * @param {CellStyleOptions} [options = {}]
 * @returns {CellStyle}
 */
export function cloneCellStyle(cellStyle, options = {}) {
    return /** @type {CellStyle} */ (mergeObjects(cellStyle, purgeObject(options)))
}
