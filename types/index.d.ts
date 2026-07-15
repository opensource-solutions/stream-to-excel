export type Color = number | string;
export type HorizontalAlignment = string;
export type VerticalAlignment = string;
export type BorderThickness = string;
export type FillPattern = string;

export type CellValue = string | number | bigint | null | undefined;

export type CellStyle = {
    type: string,
    formatCode: string,
    font: {
        name: string,
        size: number,
        color: Color | null,
        bold: boolean,
        italic: boolean,
        underline: boolean,
        strikethrough: boolean,
    },
    alignment: {
        horizontal: HorizontalAlignment | null,
        vertical: VerticalAlignment | null,
        wrapText: boolean
    },
    borderLeft: {
        thickness: BorderThickness,
        color: Color | null
    },
    borderRight: {
        thickness: BorderThickness,
        color: Color | null
    },
    borderTop: {
        thickness: BorderThickness,
        color: Color | null
    },
    borderBottom: {
        thickness: BorderThickness,
        color: Color | null
    },
    borderDiagonal: {
        thickness: BorderThickness,
        color: Color | null,
        up: boolean,
        down: boolean
    },
    fill: {
        pattern: string,
        bgColor: Color | null
    }
};

export type CellStyleOptions = {
    type?: string,
    formatCode?: string,
    font?: {
        name?: string,
        size?: number,
        color?: Color | null,
        bold?: boolean,
        italic?: boolean,
        underline?: boolean,
        strikethrough?: boolean,
    },
    alignment?: {
        horizontal?: HorizontalAlignment | null,
        vertical?: VerticalAlignment | null,
        wrapText?: boolean
    },
    borderLeft?: {
        thickness?: BorderThickness,
        color?: Color | null
    },
    borderRight?: {
        thickness?: BorderThickness,
        color?: Color | null
    },
    borderTop?: {
        thickness?: BorderThickness,
        color?: Color | null
    },
    borderBottom?: {
        thickness?: BorderThickness,
        color?: Color | null
    },
    borderDiagonal?: {
        thickness?: BorderThickness,
        color?: Color | null,
        up?: boolean,
        down?: boolean
    },
    fill?: {
        pattern?: string,
        bgColor?: Color | null
    }
};

export type Position = {
    x: number,
    y: number
};

export type Row = {
    values: CellValue[],
    styles: CellStyle[],
    doComputeExtremes?: boolean
};

export type SheetOptions = {
    author?: string,
    frozenPosition?: Position,
};

export declare function createCellStyle(options: CellStyleOptions = {}): CellStyle;
export declare function cloneCellStyle(cellStyle: CellStyle, options: CellStyleOptions = {}): CellStyle;
export declare function getColumnNameByIndex(n: number): string;
export declare function streamToExcel(stream: ReadableStream<Row>, options?: SheetOptions): Promise<Blob>;
