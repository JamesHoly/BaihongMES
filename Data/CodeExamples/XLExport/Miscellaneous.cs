using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using DevExpress.Export.Xl;
using DevExpress.XtraExport.Csv;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class Miscellaneous {

        static void Hyperlinks(Stream stream, XlDocumentFormat documentFormat) {
            #region #Hyperlinks
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 300;
                    }

                    // Local link
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Local link";
                            cell.Formatting = XlCellFormatting.Hyperlink;
                            XlHyperlink hyperlink = new XlHyperlink();
                            hyperlink.Reference = new XlCellRange(new XlCellPosition(cell.ColumnIndex, cell.RowIndex));
                            hyperlink.TargetUri = "#Sheet1!C5";
                            sheet.Hyperlinks.Add(hyperlink);
                        }
                    }

                    // External file link
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "External file link";
                            cell.Formatting = XlCellFormatting.Hyperlink;
                            XlHyperlink hyperlink = new XlHyperlink();
                            hyperlink.Reference = new XlCellRange(new XlCellPosition(cell.ColumnIndex, cell.RowIndex));
                            hyperlink.TargetUri = "linked.xlsx#Sheet1!C5";
                            sheet.Hyperlinks.Add(hyperlink);
                        }
                    }

                    // External URI
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "External uri";
                            cell.Formatting = XlCellFormatting.Hyperlink;
                            XlHyperlink hyperlink = new XlHyperlink();
                            hyperlink.Reference = new XlCellRange(new XlCellPosition(cell.ColumnIndex, cell.RowIndex));
                            hyperlink.TargetUri = "http://www.devexpress.com";
                            sheet.Hyperlinks.Add(hyperlink);
                        }
                    }
                }
            }

            #endregion #Hyperlinks
        }

        static void DocumentProperties(Stream stream, XlDocumentFormat documentFormat) {
            #region #DocumentProperties
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Setup built-in document properties
                document.Properties.Title = "Sample document";
                document.Properties.Subject = "XL Export API demo";
                document.Properties.Keywords = "XL export document generation";
                document.Properties.Description = "Generate through XL Export API";
                document.Properties.Category = "Spreadsheet";
                document.Properties.Company = "DevExpress Inc.";

                // Setup custom properties
                document.Properties.Custom["Product Suite"] = "Spreadsheet Document Automation";
                document.Properties.Custom["Revision"] = 5;
                document.Properties.Custom["Date Completed"] = DateTime.Now;
                document.Properties.Custom["Published"] = true;

                // Generate sample content
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "You can check exported document properties using File/Info/Advanced Properties dialog box.";
                        }
                    }
                }
            }

            #endregion #DocumentProperties
        }

        static void DocumentOptions(Stream stream, XlDocumentFormat documentFormat) {
            #region #DocumentOptions
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create sheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Bottom);
                    }
                    // Document format
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Document format:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.DocumentFormat.ToString().ToUpper();
                        }
                    }
                    // Max column count
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Max column count:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.MaxColumnCount;
                        }
                    }
                    // Max row count
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Max row count:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.MaxRowCount;
                        }
                    }
                    // Supports document parts
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Supports document parts:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.SupportsDocumentParts;
                        }
                    }
                    // Supports formulas
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Supports formulas:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.SupportsFormulas;
                        }
                    }
                    // Supports outline/grouping
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Supports outline/grouping:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = document.Options.SupportsOutlineGrouping;
                        }
                    }
                }
            }

            #endregion #DocumentOptions
        }

        static void CsvExportOptions(Stream stream, XlDocumentFormat documentFormat) {
            #region #CsvOptions
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Setup CSV export options
                CsvDataAwareExporterOptions csvOptions = document.Options as CsvDataAwareExporterOptions;
                if(csvOptions != null) {
                    csvOptions.Encoding = Encoding.UTF8;
                    csvOptions.WritePreamble = true;
                    csvOptions.UseCellNumberFormat = false;
                    csvOptions.NewlineAfterLastRow = true;
                }

                // Create sample content
                using(IXlSheet sheet = document.CreateSheet()) {
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                        }
                        for(int i = 0; i < 4; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Q{0}", i + 1);
                            }
                        }
                    }
                    string[] products = new string[] { 
                        "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni",
                        "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knäckebröd", "Queso Cabrales",
                        "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknödel"
                    };
                    Random random = new Random();
                    for(int i = 0; i < 12; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                            }
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                }
                            }
                        }
                    }
                }
            }

            #endregion #CsvOptions
        }

    }
}