﻿using System;
using DevExpress.Spreadsheet;
using System.Drawing;

namespace SpreadsheetExamples {
    public static class RowAndColumnActions {
        static void InsertRows(IWorkbook workbook) {
            #region #InsertRows
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with data.
            for (int i = 0; i < 10; i++) {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }

            // Insert a new row into the worksheet at the 3rd position.
            worksheet.Rows.Insert(2);

            // Insert five rows into the worksheet at the 8th position.
            worksheet.Rows.Insert(7, 5);
            #endregion #InsertRows
        }
        static void InsertColumns(IWorkbook workbook) {
            #region #InsertColumns
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with data.
            for (int i = 0; i < 10; i++) {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }
            // Insert a new column into the worksheet at the 5th position.
            worksheet.Columns.Insert(4);

            // Insert three columns into the worksheet at the 7th position.
            worksheet.Columns.Insert(6, 3);
            #endregion #InsertColumns
        }

        static void DeleteRowsColumns(IWorkbook workbook) {
            #region #DeleteRows
            Worksheet worksheet = workbook.Worksheets["Sheet1"];

            // Fill cells with data.
            for (int i = 0; i < 15; i++) {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }

            // Delete the 2nd row from the worksheet.
            worksheet.Rows.Remove(1);

            // Delete the 3rd row from the worksheet.
            worksheet.Rows[2].Delete();

            // Delete three rows from the worksheet starting from the 10th row.
            workbook.Worksheets[0].Rows.Remove(9, 3);
            #endregion #DeleteRows

            #region #DeleteColumns
            Worksheet worksheet1 = workbook.Worksheets["Sheet1"];

            // Fill cells with data.
            for (int i = 0; i < 15; i++) {
                worksheet1.Cells[i, 0].Value = i + 1;
                worksheet1.Cells[0, i].Value = i + 1;
            }
            // Delete the 2nd column from the worksheet.
            worksheet1.Columns.Remove(1);

            // Delete the 3rd column from the worksheet.
            worksheet1.Columns[2].Delete();

            // Delete three columns from the worksheet starting from the 10th column.
            workbook.Worksheets[0].Columns.Remove(9, 3);
            #endregion #DeleteColumns
        }

        static void CopyRowsColumns(IWorkbook workbook) {
            #region CopyRows

            #endregion CopyRows

            #region CopyColumns

            #endregion CopyColumns
        }

        static void ShowHideRowsColumns(IWorkbook workbook) {
            #region ShowHideRowsColumns
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the 8th row of the worksheet.
            worksheet.Rows[7].Visible = false;

            // Hide the 4th column of the worksheet.
            worksheet.Columns[3].Visible = false;

            // Populate cells with data.
            for (int i = 0; i < 10; i++) {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }
            #endregion ShowHideRowsColumns
        }

        static void SpecifyRowsHeightColumnsWidth(IWorkbook workbook) {
            #region #RowHeightAndColumnWidth
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the height of the 2nd row to 50 points
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.Rows[1].Height = 50;

            // Set the height of the row that contains the "C3" cell to 1 inches.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;
            worksheet.Cells["C3"].RowHeight = 1;

            // Set the height of the 4th row to the height of the 3rd row.
            worksheet.Rows["4"].Height = worksheet.Rows["2"].Height;

            // Set the default row height to 30 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.DefaultRowHeight = 30;

            // Set the "B" column width to 15 characters of the default font that is specified by the Normal style.
            worksheet.Columns["B"].WidthInCharacters = 15;

            // Set the "C" column width to 15 millimeters.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter;
            worksheet.Columns["C"].Width = 15;

            // Set the width of the column that contains the "E15" cell to 100 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point;
            worksheet.Cells["E15"].ColumnWidth = 100;

            // Set the width of all columns that contain the "F4:H7" cell range (the "F", "G" and "H" columns) to 70 points.
            worksheet.Range["F4:H7"].ColumnWidth = 70;

            // Set the "J" column width to the "B" column width value.
            worksheet.Columns["J"].Width = worksheet.Columns["B"].Width;

            // Copy the "C" column width value and assign it to the "K" column width.
            //worksheet.Columns["K"].CopyFrom(worksheet.Columns["C"], PasteSpecial.ColumnWidths);

            // Set the default column width to 40 pixels.
            worksheet.DefaultColumnWidthInPixels = 40;
            worksheet.Range["B1:J1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Cells["B1"].Value = "15 characters";
            worksheet.Cells["C1"].Value = "15 mm";
            worksheet.Cells["E1"].Value = "100 pt";
            worksheet.Cells["F1"].Value = "70 pt";
            worksheet.Cells["G1"].Value = "70 pt";
            worksheet.Cells["H1"].Value = "70 pt";
            worksheet.Cells["J1"].Value = "15 characters";
            //worksheet.Cells["K1"].Value = "15 mm";

            worksheet.Cells["A2"].Value = "50 pt";
            worksheet.Cells["A3"].Value = "1\"";
            worksheet.Cells["A4"].Value = "50 pt";
            Range range = worksheet.Range["A2:A5"];
            Formatting rowHeightValues = range.BeginUpdateFormatting();
            rowHeightValues.Alignment.RotationAngle = 90;
            rowHeightValues.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            rowHeightValues.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            range.EndUpdateFormatting(rowHeightValues);

            #endregion RowHeightAndColumnWidth
        }
    }
}
