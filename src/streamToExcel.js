import JSZip from 'jszip'
import * as constants from './constants.js'
import { getColumnNameByIndex } from './getColumnNameByIndex.js'
import { schedule } from './schedule.js'

/**
 * @typedef {import('../types/index.d.ts').SheetOptions} SheetOptions
 * @typedef {import('../types/index.d.ts').Row} Row
 * @typedef {import('../types/index.d.ts').CellStyle} CellStyle
 * @typedef {import('../types/index.d.ts').Color} Color
 * @typedef {{ fontId: number, alignmentId: number, borderId: number, fillId: number, numFormatId: number }} CellStyleIndexed
 */

/**
 * 
 * @param {number} fontId 
 * @param {number} alignmentId 
 * @param {number} borderId 
 * @param {number} fillId 
 * @param {number} numFormatId 
 * @returns {CellStyleIndexed}
 */
const buildStyle = (fontId, alignmentId, borderId, fillId, numFormatId) => ({ fontId, alignmentId, borderId, fillId, numFormatId })

/**
 * 
 * @param {CellStyleIndexed} a 
 * @param {CellStyleIndexed} b 
 * @returns 
 */
const stylesAreEquals = (a, b) => (
    a.fontId === b.fontId
    && a.alignmentId === b.alignmentId
    && a.borderId === b.borderId
    && a.fillId === b.fillId
    && a.numFormatId === b.numFormatId
)

/**
 * @param {unknown} value 
 * @returns {string}
 */
const xmlEscape = value => value.toString()
  .replace(/\&/g, '&amp;')
  .replace(/\</g, '&lt;')
  .replace(/\>/g, '&gt;')

/**
 * @param {string} value 
 * @param {CellStyle} style 
 * @returns {number}
 */
const length = (value, style) => {
  if (style.type === constants.TYPE_STRING)
    return value.length
  const m = style.formatCode.match(/\.(0+)$/)
  const scale = m ? m[1].length + 1 : 0
  return value.length
    + 1
    + Math.round(value.replace(/\.\d+$/, '').replace(/\D/g, '').length / 3)
    + scale
}

/**
 * @param {CellStyle} _a
 * @param {CellStyle} _b
 * @returns {boolean}
 */
const fontsAreEquals = (_a, _b) => {
    const a = _a.font
    const b = _b.font
    return a.name === b.name
        && a.size === b.size
        && a.color === b.color
        && a.bold === b.bold
        && a.italic === b.italic
        && a.underline === b.underline
        && a.strikethrough === b.strikethrough
}

/**
 * @param {CellStyle} _a 
 * @param {CellStyle} _b
 * @returns {boolean}
 */
const alignmentsAreEquals = (_a, _b) => {
    const a = _a.alignment
    const b = _b.alignment
    return a.horizontal === b.horizontal
        && a.vertical === b.vertical
        && a.wrapText === b.wrapText
}

/**
 * @param {{thickness: string, color: Color, up?: boolean, down?: boolean}} a
 * @param {{thickness: string, color: Color, up?: boolean, down?: boolean}} b 
 * @returns 
 */
const _bordersAreEquals = (a, b) => (
    a.thickness === b.thickness
        && a.color === b.color
        && a.up === b.up
        && a.down === b.down
)

/**
 * @param {CellStyle} a 
 * @param {CellStyle} b
 * @returns {boolean}
 */
const bordersAreEquals = (a, b) => {
    return _bordersAreEquals(a.borderLeft, b.borderLeft)
        && _bordersAreEquals(a.borderRight, b.borderRight)
        && _bordersAreEquals(a.borderTop, b.borderTop)
        && _bordersAreEquals(a.borderBottom, b.borderBottom)
        && _bordersAreEquals(a.borderDiagonal, b.borderDiagonal)
}

/**
 * @param {CellStyle} _a 
 * @param {CellStyle} _b
 * @returns {boolean}
 */
const fillsAreEquals = (_a, _b) => {
    const a = _a.fill
    const b = _b.fill
    return a.pattern === b.pattern
        && a.bgColor === b.bgColor
}

/**
 * @param {CellStyle} a 
 * @param {CellStyle} b
 * @returns {boolean}
 */
const numFormatAreEquals = (a, b) => {
    return a.formatCode === b.formatCode
}

/**
 * @param {CellStyle} a 
 * @returns {boolean}
 */
const hasAlignments = (a) => {
    return a.alignment.horizontal !== constants.HORIZONTAL_ALIGNMENT_DEFAULT
        || a.alignment.vertical !== constants.VERTICAL_ALIGNMENT_DEFAULT
        || a.alignment.wrapText
}

/**
 * @param {CellStyle} a 
 * @returns {boolean}
 */
const hasBorders = (a) => {
    return a.borderLeft.thickness !== constants.BORDER_THICKNESS_NONE
        || a.borderRight.thickness !== constants.BORDER_THICKNESS_NONE
        || a.borderTop.thickness !== constants.BORDER_THICKNESS_NONE
        || a.borderBottom.thickness !== constants.BORDER_THICKNESS_NONE
        || a.borderDiagonal.thickness !== constants.BORDER_THICKNESS_NONE
}

/**
 * @param {CellStyle} a 
 * @returns {boolean}
 */
const hasFills = (a) => {
    return a.fill.pattern !== constants.FILL_PATTERN_NONE
        && a.fill.bgColor !== constants.COLOR_DEFAULT
}

/**
 * @param {ReadableStream<Row>} stream
 * @param {SheetOptions} [options = {}]
 * @returns {Promise<Blob>}
 */
export async function streamToExcel(stream, options = {}) { // TODO use streams for output

    const opts = { author: '', frozenPosition: { x: 0, y: 0, ...(options.frozenPosition || {}) }, ...options }

    const { author, frozenPosition } = opts

    let rowNo = 1
    const resultRows = []
    const dimensions = {
        cols: 0,
        rows: 0
    }
    const columnsExtremes = []

    const fonts = []
    const alignments = []
    const borders = []
    const fills = []
    const styles = []
    const numFormats = []
    const strings = []

    /**
     * @param {unknown[]} haystack 
     * @param {(a: unknown, b: unknown) => boolean} comparator 
     * @returns {(target: unknown) => number}
     */
    const getId = (haystack, comparator) => target => {
        /**
         * @param {unknown} item 
         * @returns {boolean}
         */
        const criteria = item => comparator(item, target) // TODO this is slow
        let i = haystack.findIndex(criteria) // TODO this is very slow
        if (i < 0) {
            i = haystack.push(target) - 1
        }
        return i
    }

    const getFontId = getId(fonts, fontsAreEquals)
    const getAlignmentId = getId(alignments, alignmentsAreEquals)
    const getBorderId = getId(borders, bordersAreEquals)
    const getFillId = getId(fills, fillsAreEquals)
    const getNumFormatId = getId(numFormats, numFormatAreEquals)
    const getStyleId = getId(styles, stylesAreEquals)
    const getStringId = getId(strings, (a, b) => a === b)

    const iterator = stream[Symbol.asyncIterator]()

    const processRow = ({ values, styles, doComputeExtremes = true }) => {
        const row = []
        let maxFontSize = 0
        for (let colNo = 0; colNo < values.length; colNo ++) {
            const cellValue = values[colNo]
            const cellStyle = styles[colNo]
            const value = typeof cellValue === 'string' ? cellValue : (cellValue === null || cellValue === undefined ? '' : cellValue.toString())
            const fontId = getFontId(cellStyle)
            const alignmentId = getAlignmentId(cellStyle)
            const borderId = getBorderId(cellStyle)
            const fillId = getFillId(cellStyle)
            const numFormatId = getNumFormatId(cellStyle)
            const styleId = getStyleId(buildStyle(fontId, alignmentId, borderId, fillId, numFormatId))
            if (cellStyle.type === constants.TYPE_STRING) {
              const stringId = getStringId(value)
              row.push(`<c r="${ getColumnNameByIndex(colNo) + rowNo }" s="${ styleId }" t="s"><v>${ stringId }</v></c>`)
            } else {
              row.push(`<c r="${ getColumnNameByIndex(colNo) + rowNo }" s="${ styleId }" t="n"><v>${ value }</v></c>`)
            }
            if (dimensions.cols < colNo) {
                dimensions.cols = colNo
            }
            if (doComputeExtremes) {
                if (columnsExtremes[colNo] === undefined) {
                    columnsExtremes[colNo] = {
                        width: 0,
                        length: length(value, cellStyle),
                        isStyled: cellStyle.font.bold,
                        maxFontSize: cellStyle.font.size
                    }
                } else {
                    const len = length(value, cellStyle)
                    if (columnsExtremes[colNo].length < len) {
                        columnsExtremes[colNo].length = len
                        columnsExtremes[colNo].isStyled = cellStyle.font.bold
                    }
                    if (columnsExtremes[colNo].maxFontSize < cellStyle.font.size) {
                        columnsExtremes[colNo].maxFontSize = cellStyle.font.size
                    }
                }
            }
            if (maxFontSize < cellStyle.font.size) {
                maxFontSize = cellStyle.font.size
            }
        }
        if (row.length > 0) {
            resultRows.push(`<row r="${ rowNo }" spans="1:${row.length}"${ maxFontSize !== null && maxFontSize !== constants.FONT_SIZE_DEFAULT ? ` ht="${ Math.round((maxFontSize * 15.0) / constants.FONT_SIZE_DEFAULT) }" customHeight="1"` : `` }>\n${ row.join('\n') }\n</row>`)
        }
        rowNo ++
        dimensions.rows ++
    }

    let finishProcessing
    const processing = new Promise((resolve) => { finishProcessing = resolve })

    /**
     * @param {IdleDeadline} deadline 
     */
    const task = async (deadline) => {
      while (deadline.timeRemaining() > 1) {
        const { value, done } = await iterator.next()
        if (done) {
          finishProcessing()
          return
        }
        processRow(value)
      }
      schedule(task)
    }

    schedule(task)

    await processing

    columnsExtremes.forEach(col => {
        col.width = (col.length + 2)
          * (col.isStyled ? 1.1 : 1)
          * (col.maxFontSize / constants.FONT_SIZE_DEFAULT)
    })
    
    const root = new JSZip()

    const _rels = root.folder('_rels')
    const docProps = root.folder('docProps')
    const xl = root.folder('xl')
    const xl__rels = xl.folder('_rels')
    const xl_theme = xl.folder('theme')
    const xl_worksheets = xl.folder('worksheets')

    root.file('[Content_Types].xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Default Extension="xml" ContentType="application/xml" />
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
</Types>`)

    _rels.file('.rels', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
</Relationships>`)

    docProps.file('app.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>SheetJS</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Листы</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Лист1</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>12.0000</AppVersion>
</Properties>`)

    docProps.file('core.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${author}</dc:creator>
  <cp:lastModifiedBy>${author}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2022-06-21T05:48:54Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2022-06-21T05:48:54Z</dcterms:modified>
</cp:coreProperties>`)

    xl__rels.file('workbook.xml.rels', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
</Relationships>`)

    xl_theme.file('theme1.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000" />
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF" />
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D" />
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1" />
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD" />
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D" />
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59" />
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2" />
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6" />
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646" />
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF" />
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080" />
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria" />
        <a:ea typeface="" />
        <a:cs typeface="" />
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック" />
        <a:font script="Hang" typeface="맑은 고딕" />
        <a:font script="Hans" typeface="宋体" />
        <a:font script="Hant" typeface="新細明體" />
        <a:font script="Arab" typeface="Times New Roman" />
        <a:font script="Hebr" typeface="Times New Roman" />
        <a:font script="Thai" typeface="Tahoma" />
        <a:font script="Ethi" typeface="Nyala" />
        <a:font script="Beng" typeface="Vrinda" />
        <a:font script="Gujr" typeface="Shruti" />
        <a:font script="Khmr" typeface="MoolBoran" />
        <a:font script="Knda" typeface="Tunga" />
        <a:font script="Guru" typeface="Raavi" />
        <a:font script="Cans" typeface="Euphemia" />
        <a:font script="Cher" typeface="Plantagenet Cherokee" />
        <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
        <a:font script="Tibt" typeface="Microsoft Himalaya" />
        <a:font script="Thaa" typeface="MV Boli" />
        <a:font script="Deva" typeface="Mangal" />
        <a:font script="Telu" typeface="Gautami" />
        <a:font script="Taml" typeface="Latha" />
        <a:font script="Syrc" typeface="Estrangelo Edessa" />
        <a:font script="Orya" typeface="Kalinga" />
        <a:font script="Mlym" typeface="Kartika" />
        <a:font script="Laoo" typeface="DokChampa" />
        <a:font script="Sinh" typeface="Iskoola Pota" />
        <a:font script="Mong" typeface="Mongolian Baiti" />
        <a:font script="Viet" typeface="Times New Roman" />
        <a:font script="Uigh" typeface="Microsoft Uighur" />
        <a:font script="Geor" typeface="Sylfaen" />
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" />
        <a:ea typeface="" />
        <a:cs typeface="" />
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック" />
        <a:font script="Hang" typeface="맑은 고딕" />
        <a:font script="Hans" typeface="宋体" />
        <a:font script="Hant" typeface="新細明體" />
        <a:font script="Arab" typeface="Arial" />
        <a:font script="Hebr" typeface="Arial" />
        <a:font script="Thai" typeface="Tahoma" />
        <a:font script="Ethi" typeface="Nyala" />
        <a:font script="Beng" typeface="Vrinda" />
        <a:font script="Gujr" typeface="Shruti" />
        <a:font script="Khmr" typeface="DaunPenh" />
        <a:font script="Knda" typeface="Tunga" />
        <a:font script="Guru" typeface="Raavi" />
        <a:font script="Cans" typeface="Euphemia" />
        <a:font script="Cher" typeface="Plantagenet Cherokee" />
        <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
        <a:font script="Tibt" typeface="Microsoft Himalaya" />
        <a:font script="Thaa" typeface="MV Boli" />
        <a:font script="Deva" typeface="Mangal" />
        <a:font script="Telu" typeface="Gautami" />
        <a:font script="Taml" typeface="Latha" />
        <a:font script="Syrc" typeface="Estrangelo Edessa" />
        <a:font script="Orya" typeface="Kalinga" />
        <a:font script="Mlym" typeface="Kartika" />
        <a:font script="Laoo" typeface="DokChampa" />
        <a:font script="Sinh" typeface="Iskoola Pota" />
        <a:font script="Mong" typeface="Mongolian Baiti" />
        <a:font script="Viet" typeface="Arial" />
        <a:font script="Uigh" typeface="Microsoft Uighur" />
        <a:font script="Geor" typeface="Sylfaen" />
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr" />
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1" />
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="100000" />
                <a:shade val="100000" />
                <a:satMod val="130000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="50000" />
                <a:shade val="100000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0" />
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000" />
              <a:satMod val="105000" />
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr" />
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr" />
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0" />
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000" />
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400" />
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr" />
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000" />
                <a:shade val="99000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000" />
                <a:satMod val="255000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000" />
                <a:satMod val="200000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000" />
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults>
    <a:spDef>
      <a:spPr />
      <a:bodyPr />
      <a:lstStyle />
      <a:style>
        <a:lnRef idx="1">
          <a:schemeClr val="accent1" />
        </a:lnRef>
        <a:fillRef idx="3">
          <a:schemeClr val="accent1" />
        </a:fillRef>
        <a:effectRef idx="2">
          <a:schemeClr val="accent1" />
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1" />
        </a:fontRef>
      </a:style>
    </a:spDef>
    <a:lnDef>
      <a:spPr />
      <a:bodyPr />
      <a:lstStyle />
      <a:style>
        <a:lnRef idx="2">
          <a:schemeClr val="accent1" />
        </a:lnRef>
        <a:fillRef idx="0">
          <a:schemeClr val="accent1" />
        </a:fillRef>
        <a:effectRef idx="1">
          <a:schemeClr val="accent1" />
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="tx1" />
        </a:fontRef>
      </a:style>
    </a:lnDef>
  </a:objectDefaults>
  <a:extraClrSchemeLst />
</a:theme>`)

    xl.file('workbook.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506" />
  <workbookPr codeName="ThisWorkbook" />
  <bookViews>
    <workbookView xWindow="480" yWindow="855" windowWidth="28215" windowHeight="11670" />
  </bookViews>
  <sheets>
    <sheet name="Лист1" sheetId="1" r:id="rId1" />
  </sheets>
  <calcPr calcId="125725" />
</workbook>`)

    xl.file('styles.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="${ numFormats.length }">
    ${ numFormats.map((item, i) => `
      <numFmt numFmtId="${ i }" formatCode="${ item.formatCode }"/>
    `).join('\n') }
  </numFmts>
  <fonts count="${ fonts.length }">
    ${ fonts.map(item => `
      <font>
        ${ item.font.bold ? `<b />` : ``}
        ${ item.font.italic ? `<i val="true" />` : ``}
        ${ item.font.underline ? `<u val="single" />` : ``}
        ${ item.font.strikethrough ? `<strike val="true" />` : ``}
        <sz val="${ item.font.size }" />
        <color ${ item.font.color === constants.COLOR_DEFAULT ? `theme="1"` : `rgb="${item.font.color}"` } />
        <name val="${ item.font.name }" />
        <family val="2" />
        <charset val="204" />
        <scheme val="minor" />
      </font>
    `).join('\n') }
  </fonts>
  <fills count="${ fills.length }">
    ${ fills.map(item => `
      <fill>
        <patternFill patternType="${ item.fill.pattern }">
          ${ item.fill.pattern === constants.FILL_PATTERN_SOLID && item.fill.bgColor !== constants.COLOR_DEFAULT ? `<fgColor rgb="${ item.fill.bgColor }"/>` : ``}
          ${ item.fill.pattern === constants.FILL_PATTERN_SOLID && item.fill.bgColor !== constants.COLOR_DEFAULT ? `<bgColor rgb="${ item.fill.bgColor }"/>` : ``}
        </patternFill>
      </fill>
    `).join('\n') }
  </fills>
  <borders count="${ borders.length }">
    ${ borders.map(item => `
       <border${ item.borderDiagonal.thickness !== constants.BORDER_THICKNESS_NONE
          ? ` diagonalUp="${ item.borderDiagonal.up ? `true` : `false` }" diagonalDown="${ item.borderDiagonal.down ? `true` : `false` }"`
          : ``
        }>
          <left${ item.borderLeft.thickness !== constants.BORDER_THICKNESS_NONE ? ` style="${ item.borderLeft.thickness }"` : `` }>
             ${ item.borderLeft.thickness !== constants.BORDER_THICKNESS_NONE ? (item.borderLeft.color !== constants.COLOR_DEFAULT ? `<color indexed="${ item.borderLeft.color }"/>` : ``) : `` }
          </left>
          <right${ item.borderRight.thickness !== constants.BORDER_THICKNESS_NONE ? ` style="${ item.borderRight.thickness }"` : `` }>
             ${ item.borderRight.thickness !== constants.BORDER_THICKNESS_NONE ? (item.borderRight.color !== constants.COLOR_DEFAULT ? `<color indexed="${ item.borderRight.color }"/>` : ``) : `` }
          </right>
          <top${ item.borderTop.thickness !== constants.BORDER_THICKNESS_NONE ? ` style="${ item.borderTop.thickness }"` : `` }>
             ${ item.borderTop.thickness !== constants.BORDER_THICKNESS_NONE ? (item.borderTop.color !== constants.COLOR_DEFAULT ? `<color indexed="${ item.borderTop.color }"/>` : ``) : `` }
          </top>
          <bottom${ item.borderBottom.thickness !== constants.BORDER_THICKNESS_NONE ? ` style="${ item.borderBottom.thickness }"` : `` }>
             ${ item.borderBottom.thickness !== constants.BORDER_THICKNESS_NONE ? (item.borderBottom.color !== constants.COLOR_DEFAULT ? `<color indexed="${ item.borderBottom.color }"/>` : ``) : `` }
          </bottom>
          <diagonal${ item.borderDiagonal.thickness !== constants.BORDER_THICKNESS_NONE ? ` style="${ item.borderDiagonal.thickness }"` : `` }>
             ${ item.borderDiagonal.thickness !== constants.BORDER_THICKNESS_NONE ? (item.borderDiagonal.color !== constants.COLOR_DEFAULT ? `<color indexed="${ item.borderDiagonal.color }"/>` : ``) : `` }
          </diagonal>
       </border>
    `).join('\n') }
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
  </cellStyleXfs>
  <cellXfs count="${ styles.length }">
    ${ styles.map(item => {
      // const font = fonts[item.fontId]
      const alignment = alignments[item.alignmentId]
      const border = borders[item.borderId]
      const fill = fills[item.fillId]
      return `
      <xf numFmtId="${ item.numFormatId }" fontId="${ item.fontId }" fillId="${ item.fillId }" borderId="${ item.borderId }" xfId="0" applyNumberFormat="1" applyFont="1"${ hasAlignments(alignment) ? ` applyAlignment="1"` : `` }${ hasBorders(border) ? ` applyBorder="1"` : `` }${ hasFills(fill) ? ` applyFill="1"` : `` }>
        ${ hasAlignments(alignment) ? `<alignment${ alignment.alignment.horizontal !== constants.HORIZONTAL_ALIGNMENT_DEFAULT ? ` horizontal="${ alignment.alignment.horizontal }"` : `` }${ alignment.alignment.vertical !== constants.VERTICAL_ALIGNMENT_DEFAULT ? ` vertical="${ alignment.alignment.vertical }"` : `` }${ alignment.alignment.wrapText ? ` wrapText="1"` : `` } />` : `` }
      </xf>
      `
     }).join('\n') }
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Обычный" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4" />
</styleSheet>`)

    xl.file('sharedStrings.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
${ strings.map(value => `<si><t>${ xmlEscape(value) }</t></si>`).join('\n') }
</sst>`)

    xl_worksheets.file('sheet1.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${ getColumnNameByIndex(dimensions.cols) + dimensions.rows }" />
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane xSplit="${frozenPosition.x}" ySplit="${frozenPosition.y}" topLeftCell="${getColumnNameByIndex(frozenPosition.x) + (frozenPosition.y + 1)}" activePane="bottomLeft" state="frozen" />
      <selection pane="bottomLeft" activeCell="A2" sqref="A2" />
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" />
  <cols>
    ${ columnsExtremes.map((col, i) => `<col min="${ i + 1 }" max="${ i + 1 }" width="${ col.width }" customWidth="1" />`).join('\n') }
  </cols>
  <sheetData>
    ${ resultRows.join('\n') }
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>`)

    return await root.generateAsync({ type: 'blob' })
}
