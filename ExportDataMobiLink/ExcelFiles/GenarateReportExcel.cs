using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using GenaratePdf.Projection;
using OfficeOpenXml;

namespace GenaratePdf.ExcelFiles
{
    public class GenarateReportExcel
    {
        private readonly System.Drawing.Color badColor, goodColor;

        public GenarateReportExcel()
        {
            badColor = System.Drawing.Color.FromArgb(0xff, 0x90, 0x90);
            goodColor = System.Drawing.Color.FromArgb(0xb8, 0xfa, 0x9c);
        }

        private void SetRowValue(ExcelWorksheet sheet, int row, int column, object value)
        {
            sheet.Cells[row, column].Value = value;
            sheet.Cells[row, column].Style.WrapText = true;
        }

        private void SetBgColor(ExcelRange range, System.Drawing.Color color)
        {
            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        public byte[] GenerateReport(TemplatePdf report)
        {
            string templatePath = "C:/Users/GEM/source/repos/MobiLink/ExportDataMobiLink/ExcelFiles/ExportExcel/ProductReport.xlsx";
            var template = new FileInfo(templatePath);
            using (var package = new ExcelPackage(template, true))
            {
                var sheet = package.Workbook.Worksheets[1];

                sheet.Cells[2, 1].Value = string.Format(report.CompanyName);
                sheet.Cells[2, 1, 2, 2].Merge = true;
                sheet.Cells[2, 4].Value = string.Format("ĐT:" + report.CompanyPhone + "-Fax:" + report.CompanyFax);
                sheet.Cells[2, 4, 2, 7].Merge = true;
                sheet.Cells[3, 1].Value = string.Format(report.CompanyAddress);
                sheet.Cells[3, 1, 3, 2].Merge = true;
                sheet.Cells[3, 4].Value = string.Format("Tối: 38267927-39275068-39275068");
                sheet.Cells[3, 4, 3, 7].Merge = true;

                sheet.Cells[5, 1].Value = string.Format("PHIẾU BÁO GIÁ NGÀY: " + report.ReportDate);
                sheet.Cells[5, 1, 5, 7].Merge = true;
                sheet.Cells[5, 1, 5, 7].Style.Font.Bold = true;
                sheet.Cells[5, 1, 5, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[6, 1].Value = string.Format("HDL/D1/17/000008526");
                sheet.Cells[6, 1, 6, 7].Merge = true;
                sheet.Cells[6, 1, 6, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                sheet.Cells[8, 1].Value = string.Format("Khách hàng: " + report.Customer);
                sheet.Cells[8, 1, 8, 2].Merge = true;

                var row = 10;
                SetRowValue(sheet, row, 1, "TT");
                SetRowValue(sheet, row, 2, "Tên Hàng");
                SetRowValue(sheet, row, 3, "SL");
                SetRowValue(sheet, row, 4, "ĐVL");
                SetRowValue(sheet, row, 5, "Đơn Giá");
                SetRowValue(sheet, row, 6, "KM");
                SetRowValue(sheet, row, 7, "Thành Tiền");
                sheet.Cells[row, 1, row, 7].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                sheet.Cells[row, 1, row, 7].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                row = 11;
                var count = report.ListTemplate.Count;
                for (int i = 0; i < count; i++)
                {
                    var rowData = report.ListTemplate[i];
                    SetRowValue(sheet, row, 1, i + 1);
                    SetRowValue(sheet, row, 2, rowData.ProductName);
                    SetRowValue(sheet, row, 3, rowData.Number);
                    SetRowValue(sheet, row, 4, rowData.DVL);
                    SetRowValue(sheet, row, 5, rowData.Cost);
                    SetRowValue(sheet, row, 6, rowData.DiscountCost);
                    SetRowValue(sheet, row, 7, rowData.TotalCost);
                    sheet.Cells[row, 1, row, 7].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    row++;
                }
                row = row + 1;
                sheet.Cells[row, 2].Value = string.Format("Chiết khấu hóa đơn,hàng trả: 0");
                sheet.Cells[row, 2, row, 4].Merge = true;
                sheet.Cells[row, 5].Value = string.Format("Tổng cộng hóa đơn: 3,908,000");
                sheet.Cells[row, 5, row, 7].Merge = true;
                sheet.Cells[row, 5, row, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                sheet.Cells[row + 1, 5].Value = string.Format("Cộng nợ cũ: 8,127,500");
                sheet.Cells[row + 1, 5, row + 1, 7].Merge = true;
                sheet.Cells[row + 1, 5, row + 1, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                sheet.Cells[row + 2, 5].Value = string.Format("Tổng cộng tiền thanh toán: 0");
                sheet.Cells[row + 2, 5, row + 2, 7].Merge = true;
                sheet.Cells[row + 2, 5, row + 2, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                sheet.Cells[row + 3, 5].Value = string.Format("Tổng cộng nợ: 12,135,500");
                sheet.Cells[row + 3, 5, row + 3, 7].Merge = true;
                sheet.Cells[row + 3, 5, row + 3, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                row = row + 4;
                sheet.Cells[row, 1].Value = string.Format("Chú ý: Hàng mua không quá 10 ngày được trả lại");
                sheet.Cells[row, 1, row, 7].Merge = true;
                sheet.Cells[row, 1, row, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                sheet.Cells[row + 1, 2].Value = string.Format("Hàng Thượng Đình trả lại phải nguyên tem mác (Công ty yêu cầu)");
                sheet.Cells[row + 1, 2, row + 1, 7].Merge = true;
                sheet.Cells[row + 1, 2, row + 1, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;









                return package.GetAsByteArray();
            }
        }
    }
}