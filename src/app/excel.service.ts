import { Injectable } from "@angular/core";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
@Injectable({
  providedIn: "root",
})
export class ExcelService {
  constructor() {}
  STYLE_SHEET_TITLE = {
    border: false,
    height: 40,
    font: { size: 30, bold: true, color: { argb: "333333" } },
    alignment: { horizontal: "left", vertical: "middle" },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "ffffff" },
    },
  };
  STYLE_TABLE_TITLE = {
    border: true,
    height: 40,
    font: { size: 20, bold: false, color: { argb: "333333" } },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "f4f4f4" },
    },
  };
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
  STYLE_DATA_BLACK = {
    border: true,
    height: 45,
    font: { size: 15, bold: false, color: { argb: "333333" } },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "d1d1d1" },
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
  async exportToExcelHoldSections({
    myData,
    fileName,
    sheetName,
    sheetTitle,
    tableTitle,
    myHeader,
    bottomHeader,
    widths,
  }) {
    if (!myData || myData.length === 0) {
      console.error("empty data");
      return;
    }
    // console.log("exportToExcel ", myData);
    const wb = new Workbook();
    const ws = wb.addWorksheet(sheetName);
    const columns = 15;
    const data = {
      border: true,
      money: true,
      height: 0,
      font: { size: 12, bold: false, color: { argb: "000000" } },
      alignment: null,
      fill: null,
    };
    // set width từng colume
    if (widths && widths.length > 0) {
      ws.columns = widths;
    }
    // tạo ô chứa SheetTitle và merge thành một ô
    let rowSheetTitle = this.addRow(ws, [sheetTitle], this.STYLE_SHEET_TITLE);
    this.mergeRowCells(ws, rowSheetTitle, 1, columns);
    // tạo ô chứa TableTitle và merge thành một ô
    let rowTableTitle = this.addRow(ws, [tableTitle], this.STYLE_TABLE_TITLE);
    this.mergeRowCells(ws, rowTableTitle, 1, columns);

    // tạo header
    const rowHeader = this.addRow(ws, myHeader, this.STYLE_HEADER);
    rowHeader.splice(7, 0, "", "");
    rowHeader.splice(10, 0, "", "");
    this.mergeRowCells(ws, rowHeader, 6, 8);
    this.mergeRowCells(ws, rowHeader, 9, 11);
    // // header dưới
    let rowHeader2 = ws.addRow([
      "",
      "",
      "",
      "",
      "",
      ...bottomHeader,
      ...bottomHeader,
    ]);
    this.styleRowCell(rowHeader2, this.STYLE_HEADER);
    // merge dọc header
    rowHeader.eachCell({ includeEmpty: true }, function (cell, colNumber) {
      const nameOfUnderCell = `${cell._column.letter}${cell._row._number + 1}`;
      const isUnderCellHasValue = ws.getCell(nameOfUnderCell).value;
      if (!isUnderCellHasValue) {
        ws.mergeCells(`${cell._address}:${nameOfUnderCell}`);
      }
      cell;
    });
    myData.forEach((row) => {
      this.addRow(ws, row, this.STYLE_DATA_BLACK);
    });

    const buf = await wb.xlsx.writeBuffer();
    fs.saveAs(new Blob([buf]), `${fileName}.xlsx`);
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
        cell.border = borderStyles;
      }
      if (style.alignment) {
        cell.alignment = style.alignment;
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

// // Add row title and formatting
// const title = "Hole section summary";
// const titleRow = worksheet.addRow([title]);
// titleRow.font = {
//   size: 16,
//   bold: true,
// };
