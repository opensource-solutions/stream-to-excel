import * as constants from './constants.js'
import { mergeObjects } from './mergeObjects.js'
import { purgeObject } from './purgeObject.js'

/**
 * @typedef {import('../types/index.d.ts').CellStyle} CellStyle
 * @typedef {import('../types/index.d.ts').CellStyleOptions} CellStyleOptions
 */

const defaultCellStyle = {
    type: constants.TYPE_STRING,
    formatCode: constants.FORMAT_CODE_GENERAL,
    font: {
        name: constants.FONT_NAME_DEFAULT,
        size: constants.FONT_SIZE_DEFAULT,
        color: null,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
    },
    alignment: {
        horizontal: constants.HORIZONTAL_ALIGNMENT_DEFAULT,
        vertical: constants.VERTICAL_ALIGNMENT_DEFAULT,
        wrapText: false
    },
    borderLeft: {
        thickness: constants.BORDER_THICKNESS_NONE,
        color: null
    },
    borderRight: {
        thickness: constants.BORDER_THICKNESS_NONE,
        color: null
    },
    borderTop: {
        thickness: constants.BORDER_THICKNESS_NONE,
        color: null
    },
    borderBottom: {
        thickness: constants.BORDER_THICKNESS_NONE,
        color: null
    },
    borderDiagonal: {
        thickness: constants.BORDER_THICKNESS_NONE,
        color: null,
        up: true,
        down: true
    },
    fill: {
        pattern: constants.FILL_PATTERN_NONE,
        bgColor: null
    }
}

/**
 * @param {CellStyleOptions} options
 * @returns {CellStyle}
 */
export function createCellStyle(options = {}) {
    return /** @type {CellStyle} */ (mergeObjects(defaultCellStyle, purgeObject(options)))
}
