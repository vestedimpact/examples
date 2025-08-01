import { BorderStyle, ImagePosition, Workbook, Worksheet } from 'exceljs';

type CellTextProperties = {
  align: {
    horizontal: 'left' | 'center' | 'right';
    vertical: 'top' | 'middle' | 'bottom';
  };
  bold: boolean;
  color: string;
  font: string;
  italic: boolean;
  numFmt?: string;
  size: number;
  underline: boolean;
  wrapText: boolean;
};

const defaultCellText: CellTextProperties = {
  align: { horizontal: 'left', vertical: 'middle' },
  bold: false,
  color: 'FF000000',
  font: 'Arial',
  italic: false,
  size: 8,
  underline: false,
  wrapText: true,
};

const borderModes = {
  default: {
    bottom: { style: 'thin' as BorderStyle },
    left: { style: 'thin' as BorderStyle },
    right: { style: 'thin' as BorderStyle },
    top: { style: 'thin' as BorderStyle },
  },
  bottomDotted: {
    bottom: { style: 'dotted' as BorderStyle },
  },
  bottomSolid: {
    bottom: { style: 'thin' as BorderStyle },
  },
  thickVertical: {
    bottom: { style: 'thick' as BorderStyle },
    top: { style: 'thick' as BorderStyle },
  },
};

export abstract class OutputWorksheet {
  private readonly sheet: Worksheet;

  constructor(name: string, workbook: Workbook) {
    this.sheet = workbook.addWorksheet(name);
    this.sheet.getColumn('A').width = 3;
  }

  protected addBlankRows(count: number, start: number) {
    for (let i = 0; i < count; i++) {
      this.sheet.getRow(i + start).fill = {
        fgColor: { argb: 'FFFFFFFF' },
        pattern: 'solid',
        type: 'pattern',
      };
      this.sheet.getCell(`A${i + start}`).value = '';
    }
  }

  protected addBorders(cell: string, mode?: keyof typeof borderModes) {
    this.sheet.getCell(cell).border = borderModes[mode || 'default'];
  }

  protected addNumber(cell: string, content: number, overrides?: Partial<CellTextProperties>) {
    const props = { ...defaultCellText, ...(overrides || {}) };
    this.sheet.getCell(cell).alignment = { ...props.align, wrapText: props.wrapText };
    this.sheet.getCell(cell).font = {
      bold: props.bold,
      color: { argb: props.color },
      italic: props.italic,
      name: props.font,
      size: props.size,
      underline: props.underline,
    };
    this.sheet.getCell(cell).numFmt = props.numFmt || '';
    this.sheet.getCell(cell).value = content;
  }

  protected addImage(id: number, position: ImagePosition) {
    this.sheet.addImage(id, position);
  }

  protected addRichText(
    cell: string,
    items: {
      content: string;
      overrides?: Partial<CellTextProperties>;
    }[],
    align?: CellTextProperties['align'],
  ) {
    const cellAlign = align || { horizontal: 'left', vertical: 'top' };
    this.sheet.getCell(cell).alignment = { ...cellAlign, wrapText: true };
    this.sheet.getCell(cell).value = {
      richText: items.map((i) => {
        const props = { ...defaultCellText, ...(i.overrides || {}) };
        return {
          font: {
            bold: props.bold,
            color: { argb: props.color },
            italic: props.italic,
            name: props.font,
            size: props.size,
            underline: props.underline,
          },
          text: i.content,
        };
      }),
    };
  }

  protected addText(cell: string, content: string, overrides?: Partial<CellTextProperties>) {
    const props = { ...defaultCellText, ...(overrides || {}) };
    this.sheet.getCell(cell).alignment = { ...props.align, wrapText: props.wrapText };
    this.sheet.getCell(cell).font = {
      bold: props.bold,
      color: { argb: props.color },
      italic: props.italic,
      name: props.font,
      size: props.size,
      underline: props.underline,
    };
    this.sheet.getCell(cell).value = content;
  }

  protected getCell(cell: string) {
    return this.sheet.getCell(cell);
  }

  protected getColumn(col: string) {
    return this.sheet.getColumn(col);
  }

  protected getRow(row: number) {
    return this.sheet.getRow(row);
  }

  protected merge(cells: string) {
    this.sheet.mergeCells(cells);
  }

  protected setColor(cell: string, color: string) {
    this.sheet.getCell(cell).fill = {
      fgColor: { argb: color },
      pattern: 'solid',
      type: 'pattern',
    };
  }

  protected setTabColor(col: string) {
    this.sheet.properties.tabColor = { argb: col };
  }

  abstract populate(): void;
}
