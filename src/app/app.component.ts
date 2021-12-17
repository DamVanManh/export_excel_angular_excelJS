import { Component } from "@angular/core";
import { ExcelService } from "./excel.service";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  constructor(private excelService: ExcelService) {}

  generateExcel() {
    const data = [
      [
        "",
        36,
        "Conductor",
        "Casing",
        30,
        534.8,
        984.3,
        543.8,
        534.8,
        984.3,
        543.8,
        8.35,
        "",
        "",
        "Failed to load",
      ],
      [
        "",
        36,
        "Conductor",
        "Casing",
        30,
        534.8,
        984.3,
        543.8,
        534.8,
        984.3,
        543.8,
        8.35,
        "",
        "",
        "Failed to load",
      ],
      [
        "",
        36,
        "Conductor",
        "Casing",
        30,
        534.8,
        984.3,
        543.8,
        534.8,
        984.3,
        543.8,
        8.35,
        "",
        "",
        "Failed to load",
      ],
      [
        "",
        36,
        "Conductor",
        "Casing",
        30,
        534.8,
        984.3,
        543.8,
        534.8,
        984.3,
        543.8,
        8.35,
        "",
        "",
        "Failed to load",
      ],
      [
        "",
        36,
        "Conductor",
        "Casing",
        30,
        534.8,
        984.3,
        543.8,
        534.8,
        984.3,
        543.8,
        8.35,
        "",
        "",
        "Failed to load",
      ],
    ];
    const data2 = [
      [
        "Conductor casing",
        534.8,
        984.3,
        30,
        310.0,
        "X56M",
        "VIPER-3ST (M70)",
        "N/A",
        "N/A",
        "N/A",
        "N/A",
      ],
      [
        "Conductor casing",
        534.8,
        984.3,
        30,
        310.0,
        "X56M",
        "VIPER-3ST (M70)",
        "N/A",
        "N/A",
        "N/A",
        "N/A",
      ],
      [
        "Conductor casing",
        534.8,
        984.3,
        30,
        310.0,
        "X56M",
        "VIPER-3ST (M70)",
        "N/A",
        "N/A",
        "N/A",
        "N/A",
      ],
      [
        "Conductor casing",
        534.8,
        984.3,
        30,
        310.0,
        "X56M",
        "VIPER-3ST (M70)",
        "N/A",
        "N/A",
        "N/A",
        "N/A",
      ],
    ];
    const STYLE_SHEET_TITLE = {
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
    const STYLE_TABLE_TITLE = {
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
    const STYLE_DATA_BLACK = {
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
    const wb = this.excelService.generateWorkbook();
    const sheet1 = this.excelService.addWorksheet(wb, "Hole Sections");
    this.excelService.styleWidthColumns(sheet1, [
      { width: 12 },
      { width: 10 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 20 },
    ]);
    this.excelService.addRowTitle(
      sheet1,
      "Hole Sections Summary",
      1,
      15,
      STYLE_SHEET_TITLE
    );
    this.excelService.addRowTitle(
      sheet1,
      "Hole Sections",
      1,
      15,
      STYLE_TABLE_TITLE
    );
    this.excelService.addHeader(
      sheet1,
      this.excelService.STYLE_HEADER,
      [
        "",
        "Hole Size(in)",
        "Name",
        "String Type",
        "Casing OD(in)",
        "Measured Depth (ft)",
        "",
        "",
        "TVD (ft)",
        "",
        "",
        "Mud Density (ppg)",
        "Fluid Type",
        "Lithology at Shoe",
        "Kick Tolerance (bbl)",
      ],
      ["", "", "", "", "", "Top", "Shoe", "TOC", "Top", "Shoe", "TOC"]
    );
    this.excelService.addData(sheet1, data, STYLE_DATA_BLACK);

    const sheet2 = this.excelService.addWorksheet(wb, "Casing Design");
    this.excelService.styleWidthColumns(sheet2, [
      { width: 18 },
      { width: 13 },
      { width: 13 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
      { width: 18 },
    ]);
    this.excelService.addRowTitle(
      sheet2,
      "Casing Design",
      1,
      11,
      STYLE_TABLE_TITLE
    );
    this.excelService.addHeader(
      sheet2,
      this.excelService.STYLE_HEADER,
      [
        "Name/Type",
        "Top MD (ft)",
        "Base MD (ft)",
        "OD (in)",
        "Weight (ppf)",
        "Grade",
        "Connection",
        "Absolute Safety Factors",
        "",
        "",
        "",
      ],
      ["", "", "", "", "", "", "", "Burst", "Collapse", "Axial", "Triaxial"]
    );
    this.excelService.addData(sheet2, data2, STYLE_DATA_BLACK);

    this.excelService.saveAsExcelFile(wb, "Hole Sections Summary");
  }
  generateExcel2() {
    const data1 = [
      ["Jet A1", 2.7, "N/A"],
      ["Jet A1", 2.7, "N/A"],
      ["Jet A1", 2.7, "N/A"],
      ["Jet A1", 2.7, "N/A"],
      ["Jet A1", 7, "N/A"],
    ];
    const data2 = [["Temperature Gain Effect", 5.0, "N/A"]];
    const data3 = [
      ["Dealer Margin", 5.0, 0.71],
      ["Alpha", 5.0, 0.71],
    ];
    const STYLE_HEADER2 = {
      border: {
        top: { style: "thin", color: { argb: "000000" } },
        left: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "thin", color: { argb: "000000" } },
        right: { style: "thin", color: { argb: "000000" } },
      },
      height: 35,
      font: { size: 15, bold: true, color: { argb: "ffffff" } },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "579497" },
      },
    };
    const STYLE_DATA2 = {
      border: {
        top: { style: "thin", color: { argb: "000000" } },
        left: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "thin", color: { argb: "000000" } },
        right: { style: "thin", color: { argb: "000000" } },
      },
      height: 15,
      font: { size: 15, bold: true, color: { argb: "000000" } },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffffff" },
      },
    };
    const wb = this.excelService.generateWorkbook();
    const sheet1 = this.excelService.addWorksheet(wb, "Conversion factor");
    this.excelService.styleWidthColumns(sheet1, [
      { width: 50 },
      { width: 30 },
      { width: 30 },
    ]);
    this.excelService.addRowTitle(
      sheet1,
      "CONVERSION FACTOR",
      1,
      3,
      this.excelService.STYLE_HEADER
    );
    this.excelService.addHeader(sheet1, STYLE_HEADER2, [
      "CONVERSION FACTOR",
      "FACTOR",
      "USD/BBL",
    ]);
    this.excelService.addData(sheet1, data1, STYLE_DATA2);
    this.excelService.addHeader(sheet1, STYLE_HEADER2, [
      "EFFECTS",
      "FACTOR",
      "USD/BBL",
    ]);
    this.excelService.addData(sheet1, data2, STYLE_DATA2);
    this.excelService.addHeader(sheet1, STYLE_HEADER2, [
      "GASOLINE U95",
      "US CENT/LITER",
      "USD/BBL",
    ]);
    this.excelService.addData(sheet1, data3, STYLE_DATA2);
    this.excelService.saveAsExcelFile(wb, "Conversion factor");
  }
}
