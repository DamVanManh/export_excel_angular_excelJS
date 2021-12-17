import { Injectable } from "@angular/core";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
@Injectable({
  providedIn: "root",
})
export class ExcelService {
  constructor() {}
  EXCEL_TYPE =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
  EXCEL_EXTENSION = ".xlsx";

  STYLE_HEADER = {
    border: true,
    height: 35,
    font: { size: 15, bold: true, color: { argb: "000000" } },
    alignment: { horizontal: "center", vertical: "middle", wrapText: true },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "d1d1d1" },
    },
  };
  STYLE_DATA_WHITE = {
    border: true,
    height: 70,
    font: { size: 15, bold: false, color: { argb: "ffffff" } },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "ff0000" },
    },
  };

  STYLE_DATA_WARNING = {
    border: true,
    height: 70,
    font: { size: 15, bold: false, color: { argb: "ffffff" } },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "ff0000" },
    },
  };

  public addData(workSheet: any, data: any[][], style: any) {
    data.forEach((row) => {
      this.addRow(workSheet, row, style);
    });
  }
  public addHeader(
    ws: any,
    style: any,
    header: string[],
    bottomHeader?: string[]
  ) {
    const rowHeader = this.addRow(ws, header, style);
    let rowBottomHeader;
    // merge empty cell horizontal
    for (let indexContent = 0; indexContent < header.length; indexContent++) {
      if (header[indexContent] === "" && indexContent > 0) {
        const cellFrom = indexContent;
        let cellTo = 0;
        for (let index = indexContent + 1; index < header.length + 1; index++) {
          console.log(
            "vào tới ",
            indexContent,
            index,
            header[index],
            header.length
          );

          if (header[index] !== "" || index === header.length) {
            cellTo = index;
            indexContent = index;
            this.mergeRowCells(ws, rowHeader, cellFrom, cellTo);
            break;
          }
        }
      }
    }
    // merge empty cell vertical
    if (bottomHeader && bottomHeader.length > 0) {
      rowBottomHeader = this.addRow(ws, bottomHeader, style);
      rowHeader.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        const nameOfUnderCell = `${cell._column.letter}${
          cell._row._number + 1
        }`;
        const isUnderCellHasValue = ws.getCell(nameOfUnderCell).value;
        if (!isUnderCellHasValue) {
          ws.mergeCells(`${cell._address}:${nameOfUnderCell}`);
        }
        cell;
      });
    }
    return { rowHeader, rowBottomHeader };
  }
  private addRow(ws, data, style) {
    const row = ws.addRow(data);
    this.styleRowCell(row, style);
    return row;
  }
  private styleRowCell(row, style) {
    const borderStyles = {
      top: { style: "thin", color: { argb: "858585" } },
      left: { style: "thin", color: { argb: "858585" } },
      bottom: { style: "thin", color: { argb: "858585" } },
      right: { style: "thin", color: { argb: "858585" } },
    };
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (style.border) {
        cell.border = { ...borderStyles, ...style.border };
      }
      if (style.alignment) {
        cell.alignment = { ...style.alignment, wrapText: true };
      } else {
        cell.alignment = { vertical: "middle" };
      }
      if (style.font) {
        cell.font = style.font;
      }
      if (style.fill) {
        cell.fill = style.fill;
      }
    });
    if (style.height > 0) {
      row.height = style.height;
    }
  }
  private mergeRowCells(ws, row, from, to) {
    ws.mergeCells(`${row.getCell(from)._address}:${row.getCell(to)._address}`);
  }
  public async saveAsExcelFile(wb: any, fileName: string): Promise<void> {
    const buffer = await wb.xlsx.writeBuffer();
    const data: Blob = new Blob([buffer], {
      type: this.EXCEL_TYPE,
    });
    fs.saveAs(data, fileName + new Date().getTime() + this.EXCEL_EXTENSION);
  }
  public generateWorkbook(): Workbook {
    return new Workbook();
  }
  public addWorksheet(wb: Workbook, sheetName: string): any {
    return wb.addWorksheet(sheetName);
  }
  public addRowTitle(
    workSheet: any,
    title: string,
    from: number,
    to: number,
    style: any
  ): any {
    let rowSheetTitle = this.addRow(workSheet, [title], style);
    this.mergeRowCells(workSheet, rowSheetTitle, from, to);
  }
  public styleWidthColumns(workSheet: any, widths: { width: number }[]) {
    if (widths && widths.length > 0) {
      workSheet.columns = widths;
    }
  }
  private addEmptyRow(workSheet: any, numberRow: number = 1) {
    for (let index = 0; index < numberRow; index++) {
      workSheet.addRow([]);
    }
  }
}

// data.forEach((d) => {
//   const row = worksheet.addRow(d);
//   const qty = row.getCell(5);
//   let color = "FF99FF99";
//   if (+qty.value < 500) {
//     color = "FF9999";
//   }

//   qty.fill = {
//     type: "pattern",
//     pattern: "solid",
//     fgColor: { argb: color },
//   };
// });

// worksheet.getColumn(3).width = 30;
// worksheet.getColumn(4).width = 30;
// rowHeader.splice(7, 0, "", "");
// // Add row title and formatting
// const title = "Hole section summary";
// const titleRow = worksheet.addRow([title]);
// titleRow.font = {
//   size: 16,
//   bold: true,
// };
