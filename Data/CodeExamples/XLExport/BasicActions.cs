using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class BasicActions {

        static void CreateDocument(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateDocument
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            
            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;
            }

            #endregion #CreateDocument
        }

        static void CreateSheet(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateSheet
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            
            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.Name = "Sales report";
                }
            }

            #endregion #CreateSheet
        }

        static void CreateHiddenSheet(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateHiddenSheet
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            
            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.Name = "Sales report";
                }

                // Create worksheet and setup visibility
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.Name = "Data";
                    sheet.VisibleState = XlSheetVisibleState.Hidden;
                }
            }

            #endregion #CreateHiddenSheet
        }

        static void CreateColumns(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateColumns
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            
            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Set column A width to 100 pixels
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                    }

                    // Hide column B
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.IsHidden = true;
                    }

                    // Set column D width of 24.5 characters
                    using(IXlColumn column = sheet.CreateColumn(3)) {
                        column.WidthInCharacters = 24.5f;
                    }
                }
            }

            #endregion #CreateColumns
        }

        static void CreateRows(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateRows
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Set row 1 height to 40 pixels
                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 40;
                    }

                    // Hide row 3
                    using(IXlRow row = sheet.CreateRow(2)) {
                        row.IsHidden = true;
                    }
                }
            }

            #endregion #CreateRows
        }

        static void CreateCells(Stream stream, XlDocumentFormat documentFormat) {
            #region #CreateCells
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {

                // Setup document culture
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 150;
                    }

                    using(IXlRow row = sheet.CreateRow()) {

                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Numeric value:";
                        }

                        // Create cell with numeric value
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 123.45;
                        }
                    }
                    
                    using(IXlRow row = sheet.CreateRow()) {

                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Text value:";
                        }

                        // Create cell with text value
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "abc";
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {

                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Boolean value:";
                        }

                        // Create cell with boolean value
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = true;
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {

                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Error value:";
                        }

                        // Create cell with error value
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = XlVariantValue.ErrorName;
                        }
                    }
                }
            }

            #endregion #CreateCells
        }

        static void MergeCells(Stream stream, XlDocumentFormat documentFormat) {
            #region #MergeCells
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                using(IXlSheet sheet = document.CreateSheet()) {
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Merged cells in range A1:E1";
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Merged cells in range A2:A5";
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                            cell.Formatting.Alignment.WrapText = true;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Merged cells in range B2:E5";
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                        }
                    }

                    // Merge cells in range A1:E1
                    sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 0, 4, 0));

                    // Merge cells in range A2:A5
                    sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 1, 0, 4));

                    // Merge cells in range B2:E5
                    sheet.MergedCells.Add(XlCellRange.FromLTRB(1, 1, 4, 4));
                }
            }

            #endregion #MergeCells
        }

    }
}