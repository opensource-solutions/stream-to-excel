import { writeFileSync } from 'fs'
import {
    createCellStyle,
    cloneCellStyle,
    streamToExcel,
    BORDER_THICKNESS_THIN, TYPE_NUMERIC, HORIZONTAL_ALIGNMENT_RIGHT
} from "../dist/stream-to-excel.min.js"

const createExampleStream = () => {
    let isCanceled = false
    let numRows = 10
    const status = ['Нет связи','Готов','Выполняется']
    const style = createCellStyle({
        borderLeft: { thickness: BORDER_THICKNESS_THIN },
        borderRight: { thickness: BORDER_THICKNESS_THIN },
        borderBottom: { thickness: BORDER_THICKNESS_THIN },
        borderTop: { thickness: BORDER_THICKNESS_THIN },
    })
    const totals = [ 0, 0, 0 ]
    return new ReadableStream({
        start(controller) {
            controller.enqueue({
                values: [
                    'Какие-то интересные данные для примера'
                ],
                styles: [
                    createCellStyle()
                ]
            })
            controller.enqueue({
                values: [
                    'Идентификатор',
                    'IP-адрес',
                    'Статус',
                    'Вх.трафик',
                    'Исх.трафик',
                    'Заданий',
                ],
                styles: [
                    cloneCellStyle(style, { font: { bold: true }, }),
                    cloneCellStyle(style, { font: { bold: true }, }),
                    cloneCellStyle(style, { font: { bold: true }, }),
                    cloneCellStyle(style, { font: { bold: true }, alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                    cloneCellStyle(style, { font: { bold: true }, alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                    cloneCellStyle(style, { font: { bold: true }, alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                ]
            })
        },
        pull(controller) {
            if (isCanceled || numRows <= 0) {
                if (!isCanceled) {
                    controller.enqueue({
                        values: [
                            'Итого',
                            '',
                            '',
                            totals[0],
                            totals[1],
                            totals[2],
                        ],
                        styles: [
                            cloneCellStyle(style),
                            cloneCellStyle(style),
                            cloneCellStyle(style),
                            cloneCellStyle(style, { type: TYPE_NUMERIC, formatCode: '#,##0.000', alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                            cloneCellStyle(style, { type: TYPE_NUMERIC, formatCode: '#,##0.000', alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                            cloneCellStyle(style, { type: TYPE_NUMERIC, alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                        ]
                    })
                }
                controller.close()
                return
            }
            const values = [
                crypto.randomUUID(),
                `192.168.20.${Math.floor(Math.random() * 200) + 10}`,
                status.at(Math.floor(Math.random() * status.length)),
                Math.floor(Math.random() * 5000000),
                Math.floor(Math.random() * 200000),
                Math.floor(Math.random() * 20),
            ]
            totals[0] += /** @type {number} */ (values[3])
            totals[1] += /** @type {number} */ (values[4])
            totals[2] += /** @type {number} */ (values[5])
            controller.enqueue({
                values,
                styles: [
                    cloneCellStyle(style),
                    cloneCellStyle(style),
                    cloneCellStyle(style),
                    cloneCellStyle(style, { type: TYPE_NUMERIC, formatCode: '#,##0.000', alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                    cloneCellStyle(style, { type: TYPE_NUMERIC, formatCode: '#,##0.000', alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                    cloneCellStyle(style, { type: TYPE_NUMERIC, alignment: { horizontal: HORIZONTAL_ALIGNMENT_RIGHT } }),
                ]
            })
            numRows --
        },
        cancel() {
            isCanceled = true
        }
    })
}

streamToExcel(createExampleStream(), { author: 'Example', frozenPosition: { x: 0, y: 2 } })
    .then(async (blob) => {
        writeFileSync('./Пример выгрузки.xlsx', Buffer.from(await blob.arrayBuffer()))
    })
    .catch(error => {
        console.error(error)
        globalThis.process.exit(1)
    })
