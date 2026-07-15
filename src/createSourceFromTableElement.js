import { createCellStyle } from './createCellStyle.js'
import * as constants from './constants.js'

/**
 * @typedef {import('../types/index.d.ts').Row} Row
 * @typedef {import('../types/index.d.ts').TableToSourceOptions} TableToSourceOptions
 * @typedef {import('../types/index.d.ts').WorksheetSource} WorksheetSource
 */

/**
 * @param {HTMLTableElement} tableElement 
 * @param {TableToSourceOptions} [options = {}]
 * @returns {WorksheetSource}
 */
export function createSourceFromTableElement(tableElement, options = {}) {

    const { skipEmptyRows = true, author = '' } = options

    const sourceRows = [
        ...Array.from([ tableElement ], (rowElement) => ({ rowElement, cellsSelector: 'caption', skipEmptyRows: true, doComputeExtremes: false, isHeader: true })),
        ...Array.from(tableElement.querySelectorAll('thead > tr'), (rowElement) => ({ rowElement, cellsSelector: 'th, td', skipEmptyRows, doComputeExtremes: true, isHeader: true })),
        ...Array.from(tableElement.querySelectorAll('tbody > tr'), (rowElement) => ({ rowElement, cellsSelector: 'td', skipEmptyRows, doComputeExtremes: true, isHeader: false })),
        ...Array.from(tableElement.querySelectorAll('tfoot > tr'), (rowElement) => ({ rowElement, cellsSelector: 'td', skipEmptyRows, doComputeExtremes: true, isHeader: false })),
    ]

    /**
     * @param {string|undefined} x 
     * @returns {boolean}
     */
    const toBoolOrUndefined = x => x === undefined ? undefined : /^(true|yes|on|1)$/i.test(x)

    return {

        async getAuthor() {
            return author
        },

        async getFrozenPosition() {
            const numRows = sourceRows.reduce((result, { isHeader }) => isHeader ? result + 1 : result, 0)
            return { x: 0, y: numRows }
        },

        async getReadableStream() {
            let rowIndex = 0
            let isCanceled = false
            return new ReadableStream({
                async pull(controller) {
                    if (isCanceled) {
                        controller.close()
                        return
                    }
                    for (;;) {
                        if (rowIndex >= sourceRows.length) {
                            controller.close()
                            return
                        }
                        const { rowElement, cellsSelector, skipEmptyRows, doComputeExtremes } = sourceRows[rowIndex ++]
                        const values = []
                        const styles = []
                        const cells = rowElement.querySelectorAll(cellsSelector)
                        for (let i = 0; i < cells.length; i ++) {
                            const cellElement = /** @type {HTMLElement} */ (cells[i])
                            const opts = cellElement.dataset
                            const isHidden = opts.xlHidden || false
                            if (isHidden) continue
                            const cellStyle = createCellStyle({
                                type: (/^(numeric|number|num|n)$/i.test(opts.xlType)
                                    ? constants.TYPE_NUMERIC
                                    : constants.TYPE_STRING),
                                formatCode: opts.xlFmt,
                                font: {
                                    name: opts.xlFont,
                                    size: opts.xlFontSize ? Number(opts.xlFontSize) : undefined,
                                    color: opts.xlColor,
                                    bold: toBoolOrUndefined(opts.xlBold),
                                    italic: toBoolOrUndefined(opts.xlItalic),
                                    underline: toBoolOrUndefined(opts.xlUnderline),
                                    strikethrough: toBoolOrUndefined(opts.xlStrikethrough),
                                },
                                alignment: {
                                    horizontal: opts.xlHalign,
                                    vertical: opts.xlValign,
                                    wrapText: toBoolOrUndefined(opts.xlWrap),
                                },
                                borderLeft: {
                                    thickness: opts.xlBorderLeft || opts.xlBorder,
                                },
                                borderRight: {
                                    thickness: opts.xlBorderLeft || opts.xlBorder,
                                },
                                borderBottom: {
                                    thickness: opts.xlBorderLeft || opts.xlBorder,
                                },
                                borderTop: {
                                    thickness: opts.xlBorderLeft || opts.xlBorder,
                                },
                                borderDiagonal: {
                                    thickness: opts.xlBorderDiagonal,
                                    up: true,
                                    down: true,
                                },
                                fill: {
                                    pattern: opts.xlBackgroundColor
                                        ? constants.FILL_PATTERN_SOLID
                                        : constants.FILL_PATTERN_NONE,
                                    bgColor: opts.xlBackgroundColor,
                                }
                            })
                            styles.push(cellStyle)
                            values.push(cellElement.innerText)
                        }
                        if (skipEmptyRows && values.length <= 0)
                            continue
                        controller.enqueue({ values, styles, doComputeExtremes })
                        break
                    }
                },
                cancel() {
                    isCanceled = true
                }
            })
        }
    }
}
