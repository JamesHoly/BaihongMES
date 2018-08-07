using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class DataActions {

        static void AutoFilter(Stream stream, XlDocumentFormat documentFormat) {
            #region #AutoFilter
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create columns
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 250;
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.Formatting = new XlCellFormatting();
                        column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                    }

                    // Prepare cells formatting
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Prepare header row formatting
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));

                    // Generate header row
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Region";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Product";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Sales";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate sample content
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };
                    int[] amount = new int[] { 6750, 4500, 3550, 4250, 5500, 6250, 5325, 4235 };
                    for(int i = 0; i < 8; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = (i < 4) ? "East" : "West";
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i % 4];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = amount[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }

                    // Set auto filter
                    sheet.AutoFilterRange = sheet.DataRange;
                }
            }

            #endregion #AutoFilter
        }

        static void OutlineGrouping(Stream stream, XlDocumentFormat documentFormat) {
            #region #Group/Outline
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Setup outline properties
                    sheet.OutlineProperties.SummaryBelow = true;
                    sheet.OutlineProperties.SummaryRight = true;

                    // Create Region/Product column
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 200;
                    }
                    // Create and group Q1-Q4 columns
                    sheet.BeginGroup(false);
                    for(int i = 0; i < 4; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }
                    sheet.EndGroup();
                    // Create Yearly total column
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.Formatting = new XlCellFormatting();
                        column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                    }

                    // Prepare cells formatting
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = XlFont.BodyFont();
                    rowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light1, 0.0));
                    // Prepare header row formatting
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.Font = XlFont.BodyFont();
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent1, 0.0));
                    // Prepare total row formatting
                    XlCellFormatting totalRowFormatting = new XlCellFormatting();
                    totalRowFormatting.Font = XlFont.BodyFont();
                    totalRowFormatting.Font.Bold = true;
                    totalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0));
                    // Prepare grand total row formatting
                    XlCellFormatting grandTotalRowFormatting = new XlCellFormatting();
                    grandTotalRowFormatting.Font = XlFont.BodyFont();
                    grandTotalRowFormatting.Font.Bold = true;
                    grandTotalRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, -0.2));

                    // Generate sample content
                    Random random = new Random();
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni" };

                    // Group rows for grand total
                    sheet.BeginGroup(false);
                    for(int p = 0; p < 2; p++) {
                        // Generate header row
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = (p == 0) ? "East" : "West";
                                cell.ApplyFormatting(headerRowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));
                            }
                            for(int i = 0; i < 4; i++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = string.Format("Q{0}", i + 1);
                                    cell.ApplyFormatting(headerRowFormatting);
                                    cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                                }
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = "Yearly total";
                                cell.ApplyFormatting(headerRowFormatting);
                                cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                            }
                        }

                        // Create and group data rows
                        sheet.BeginGroup(false);
                        for(int i = 0; i < 4; i++) {
                            using(IXlRow row = sheet.CreateRow()) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = products[i];
                                    cell.ApplyFormatting(rowFormatting);
                                    cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.8));
                                }
                                for(int j = 0; j < 4; j++) {
                                    using(IXlCell cell = row.CreateCell()) {
                                        cell.Value = Math.Round(random.NextDouble() * 2000 + 3000);
                                        cell.ApplyFormatting(rowFormatting);
                                    }
                                }
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.SetFormula(XlFunc.Sum(XlCellRange.FromLTRB(1, row.RowIndex, 4, row.RowIndex)));
                                    cell.ApplyFormatting(rowFormatting);
                                    cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                                }
                            }
                        }
                        sheet.EndGroup();

                        // Create total row
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = "Total";
                                cell.ApplyFormatting(totalRowFormatting);
                                cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.6));
                            }
                            for(int j = 0; j < 5; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, row.RowIndex - 4, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                    cell.ApplyFormatting(totalRowFormatting);
                                }
                            }
                        }
                    }
                    sheet.EndGroup();

                    // Create grand total row
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Grand total";
                            cell.ApplyFormatting(grandTotalRowFormatting);
                            cell.Formatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.4));
                        }
                        for(int j = 0; j < 5; j++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.SetFormula(XlFunc.Subtotal(XlCellRange.FromLTRB(j + 1, 1, j + 1, row.RowIndex - 1), XlSummary.Sum, false));
                                cell.ApplyFormatting(grandTotalRowFormatting);
                            }
                        }
                    }
                }
            }

            #endregion #Group/Outline
        }

        static void DataValidations(Stream stream, XlDocumentFormat documentFormat) {
            #region #DataValidation
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create columns
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 110;
                        column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom);
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 190;
                    }
                    for(int i = 0; i < 2; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 90;
                            column.Formatting = new XlCellFormatting();
                            column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                    }
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 130;
                    }

                    sheet.SkipColumns(1);
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 130;
                    }

                    // Prepare cells formatting
                    XlCellFormatting rowFormatting = new XlCellFormatting();
                    rowFormatting.Font = new XlFont();
                    rowFormatting.Font.Name = "Century Gothic";
                    rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None;

                    // Prepare header row formatting
                    XlCellFormatting headerRowFormatting = new XlCellFormatting();
                    headerRowFormatting.CopyFrom(rowFormatting);
                    headerRowFormatting.Font.Bold = true;
                    headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0));

                    // Generate header row
                    using(IXlRow row = sheet.CreateRow()) {
                        string[] columnNames = new string[] { "Employee ID", "Employee name", "Salary", "Bonus", "Department" };
                        row.BulkCells(columnNames, headerRowFormatting);
                        row.SkipCells(1);
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Departments";
                            cell.ApplyFormatting(headerRowFormatting);
                        }
                    }

                    // Generate sample content
                    int[] id = new int[] {10115, 10709, 10401, 10204 };
                    string[] name = new string[] { "Augusta Delono", "Chris Cadwell", "Frank Diamond", "Simon Newman" };
                    int[] salary = new int[] { 1100, 2000, 1750, 1250 };
                    int[] bonus = new int[] { 50, 180, 100, 80 };
                    int[] deptid = new int[] { 0, 2, 3, 3 };
                    string[] department = new string[] { "Accounting", "IT", "Management", "Manufacturing" };
                    for(int i = 0; i < 4; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = id[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = name[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = salary[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = bonus[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = department[deptid[i]];
                                cell.ApplyFormatting(rowFormatting);
                            }
                            row.SkipCells(1);
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = department[i];
                                cell.ApplyFormatting(rowFormatting);
                            }
                        }
                    }

                    // Setup data validations

                    // Custom data validation (Employee ID must be a 5-digit number)
                    XlDataValidation validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(0, 1, 0, 4)); // A2:A5
                    validation.Type = XlDataValidationType.Custom;
                    validation.Criteria1 = "=AND(ISNUMBER(A2),LEN(A2)=5)";
                    sheet.DataValidations.Add(validation);

                    // Whole number data validation with prompt and warning message (Salary must be in range $600...$2000)
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(2, 1, 2, 4)); // C2:C5
                    validation.Type = XlDataValidationType.Whole;
                    validation.Operator = XlDataValidationOperator.Between;
                    validation.Criteria1 = 600;
                    validation.Criteria2 = 2000;
                    validation.ErrorMessage = "Salary can be set in the range $600-$2000.";
                    validation.ErrorTitle = "Warning";
                    validation.ErrorStyle = XlDataValidationErrorStyle.Warning;
                    validation.InputPrompt = "Please enter whole number in range 600...2000";
                    validation.PromptTitle = "Salary";
                    validation.ShowErrorMessage = true;
                    validation.ShowInputMessage = true;
                    sheet.DataValidations.Add(validation);

                    // Decimal value data validation with informational message (Bonus cannot be greater than 10% of the salary.)
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(3, 1, 3, 4)); // D2:D5
                    validation.Type = XlDataValidationType.Whole;
                    validation.Operator = XlDataValidationOperator.Between;
                    validation.Criteria1 = 0;
                    validation.Criteria2 = "=C2*0.1";
                    validation.ErrorMessage = "Bonus cannot be greater than 10% of the salary.";
                    validation.ErrorTitle = "Information";
                    validation.ErrorStyle = XlDataValidationErrorStyle.Information;
                    validation.ShowErrorMessage = true;
                    sheet.DataValidations.Add(validation);

                    // List data validation (Department must be one of the values from the Departments list)
                    validation = new XlDataValidation();
                    validation.Ranges.Add(XlCellRange.FromLTRB(4, 1, 4, 4)); // E2:E5
                    validation.Type = XlDataValidationType.List;
                    validation.Criteria1 = XlCellRange.FromLTRB(6, 1, 6, 4).AsAbsolute(); // $G$2:$G$5
                    sheet.DataValidations.Add(validation);
                }
            }

            #endregion #DataValidation
        }

    }
}