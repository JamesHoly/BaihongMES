using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class ConditionalFormatting {

        static void Average(Stream stream, XlDocumentFormat documentFormat) {
            #region #AverageRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = i + 1;
                                }
                            }
                        }
                    }

                    // Highlight values above average
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); // A1:A11
                    XlCondFmtRuleAboveAverage rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.Above;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    // Highlight values above or equal to average
                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)); // B1:B11
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.AboveOrEqual;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    // Highlight values below average
                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)); // C1:C11
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.Below;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    // Highlight values below or equal to average
                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)); // D1:D11
                    rule = new XlCondFmtRuleAboveAverage();
                    rule.Condition = XlCondFmtAverageCondition.BelowOrEqual;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #AverageRule
        }

        static void CellIs(Stream stream, XlDocumentFormat documentFormat) {
            #region #CellIsRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = i + 1;
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = 12 - i;
                            }
                        }
                    }

                    // Highlight values using "cell is" rules with "less than", "between" and "greater than" operators
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); // A1:A11
                    XlCondFmtRuleCellIs rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.LessThan;
                    rule.Formatting = XlCellFormatting.Bad;
                    rule.Value = 5;
                    formatting.Rules.Add(rule);
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.Between;
                    rule.Formatting = XlCellFormatting.Neutral;
                    rule.Value = 5;
                    rule.SecondValue = 8;
                    formatting.Rules.Add(rule);
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.GreaterThan;
                    rule.Formatting = XlCellFormatting.Good;
                    rule.Value = 8;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    // Highlight values using "cell is" rule with criteria value specified by formula
                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)); // B1:B11
                    rule = new XlCondFmtRuleCellIs();
                    rule.Operator = XlCondFmtOperator.GreaterThan;
                    rule.Formatting = XlCellFormatting.Bad;
                    rule.Value = "=$A1+3";
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #CellIsRule
        }

        static void Blanks(Stream stream, XlDocumentFormat documentFormat) {
            #region #BlanksRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 10; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                if((i % 2) == 0)
                                    cell.Value = i + 1;
                            }
                        }
                    }

                    // Format cells with blanks/no blanks
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9)); // A1:A10
                    // Highlight blank cells
                    XlCondFmtRuleBlanks rule = new XlCondFmtRuleBlanks(true);
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Highlight non blank cells
                    rule = new XlCondFmtRuleBlanks(false); // non blank cells
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #BlanksRule
        }

        static void Duplicates(Stream stream, XlDocumentFormat documentFormat) {
            #region #DuplicatesRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.ColumnIndex * cell.RowIndex + cell.RowIndex + 1;
                                }
                            }
                        }
                    }

                    // Highlight duplicate/unique values
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 10)); // A1:D11
                    // Highlight duplicate values
                    formatting.Rules.Add(new XlCondFmtRuleDuplicates() { Formatting = XlCellFormatting.Bad });
                    // Highlight unique values
                    formatting.Rules.Add(new XlCondFmtRuleUnique() { Formatting = XlCellFormatting.Good });
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #DuplicatesRule
        }

        static void Expression(Stream stream, XlDocumentFormat documentFormat) {
            #region #ExpressionRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat, new XlFormulaParser());

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    int[] width = new int[] { 80, 150, 90 };
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = width[i];
                            if(i == 2) {
                                column.Formatting = new XlCellFormatting();
                                column.Formatting.NumberFormat = "[$$-409] #,##0.00";
                            }
                        }
                    }
                    string[] columnNames = new string[] { "Account ID", "User Name", "Balance" };
                    using(IXlRow row = sheet.CreateRow()) {
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.Font = XlFont.BodyFont();
                        headerRowFormatting.Font.Bold = true;
                        headerRowFormatting.Border = new XlBorder();
                        headerRowFormatting.Border.BottomColor = Color.Black;
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin;
                        for(int i = 0; i < 3; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = columnNames[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }
                    string[] accountIds = new string[] { "A105", "A114", "B013", "C231", "D101", "D105" };
                    string[] users = new string[] { "Berry Dafoe", "Chris Cadwell", "Esta Mangold", "Liam Bell", "Simon Newman", "Wendy Underwood" };
                    int[] balance = new int[] { 155, 250, 48, 350, -15, 10 };
                    for(int i = 0; i < 6; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = accountIds[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = users[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = balance[i];
                            }
                        }
                    }

                    // Highlight values using expression rule
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 1, 2, 6)); // A2:C7
                    XlCondFmtRuleExpression rule = new XlCondFmtRuleExpression("AND($C2>0,$C2<50)");
                    rule.Formatting = XlFill.SolidFill(Color.FromArgb(0xff, 0xff, 0xcc));
                    formatting.Rules.Add(rule);
                    rule = new XlCondFmtRuleExpression("$C2<=0");
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #ExpressionRule
        }

        static void SpecificText(Stream stream, XlDocumentFormat documentFormat) {
            #region #SpecificTextRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    int[] width = new int[] { 250, 180, 100 };
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = width[i];
                            if(i == 2) {
                                column.Formatting = new XlCellFormatting();
                                column.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                            }
                        }
                    }
                    string[] columnNames = new string[] { "Product", "Delivery", "Sales" };
                    using(IXlRow row = sheet.CreateRow()) {
                        XlCellFormatting headerRowFormatting = new XlCellFormatting();
                        headerRowFormatting.Font = XlFont.BodyFont();
                        headerRowFormatting.Font.Bold = true;
                        headerRowFormatting.Border = new XlBorder();
                        headerRowFormatting.Border.BottomColor = Color.Black;
                        headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin;
                        for(int i = 0; i < 3; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = columnNames[i];
                                cell.ApplyFormatting(headerRowFormatting);
                            }
                        }
                    }
                    string[] products = new string[] { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Queso Cabrales", "Raclette Courdavault" };
                    string[] deliveries = new string[] { "USA", "Worldwide", "USA", "Ships worldwide", "Worldwide except EU", "EU" };
                    int[] sales = new int[] { 15500, 20250, 12634, 35010, 15234, 10050 };
                    for(int i = 0; i < 6; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = products[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = deliveries[i];
                            }
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = sales[i];
                            }
                        }
                    }

                    // Highlight values using specific text rule
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 1, 1, 6)); // B2:B7
                    XlCondFmtRuleSpecificText rule = new XlCondFmtRuleSpecificText(XlCondFmtSpecificTextType.Contains, "worldwide");
                    rule.Formatting = XlCellFormatting.Neutral;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #SpecificTextRule
        }

        static void TimePeriod(Stream stream, XlDocumentFormat documentFormat) {
            #region #TimePeriodRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 100;
                        column.ApplyFormatting(XlNumberFormat.ShortDate);
                    }
                    for(int i = 0; i < 10; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = DateTime.Now.AddDays(row.RowIndex - 5);
                            }
                        }
                    }

                    // Highlight values using time period rules
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9)); // A1:A10
                    XlCondFmtRuleTimePeriod rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Yesterday;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Today;
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    rule = new XlCondFmtRuleTimePeriod();
                    rule.TimePeriod = XlCondFmtTimePeriod.Tomorrow;
                    rule.Formatting = XlCellFormatting.Neutral;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #TimePeriodRule
        }

        static void Top10(Stream stream, XlDocumentFormat documentFormat) {
            #region #Top/BottomRules
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 10; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.ColumnIndex * 4 + cell.RowIndex + 1;
                                }
                            }
                        }
                    }

                    // Highlight values using top/bottom rules
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 9)); // A1:D10
                    // Bottom 10
                    XlCondFmtRuleTop10 rule = new XlCondFmtRuleTop10();
                    rule.Bottom = true;
                    rule.Formatting = XlCellFormatting.Bad;
                    formatting.Rules.Add(rule);
                    // Top 10
                    rule = new XlCondFmtRuleTop10();
                    rule.Formatting = XlCellFormatting.Good;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #Top/BottomRules
        }

        static void DataBar(Stream stream, XlDocumentFormat documentFormat) {
            #region #DataBarRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 3; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    int rowIndex = cell.RowIndex;
                                    int columnIndex = cell.ColumnIndex;
                                    if(columnIndex == 0)
                                        cell.Value = rowIndex + 1;
                                    else if(columnIndex == 1)
                                        cell.Value = rowIndex - 5;
                                    else
                                        cell.Value = (rowIndex < 5) ? rowIndex + 1 : 11 - rowIndex;
                                }
                            }
                        }
                    }

                    // Format values using data bar rule
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); // A1:A11
                    XlCondFmtRuleDataBar rule = new XlCondFmtRuleDataBar();
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.2);
                    rule.GradientFill = false;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)); // B1:B11
                    rule = new XlCondFmtRuleDataBar();
                    rule.FillColor = Color.Green;
                    rule.BorderColor = Color.Green;
                    rule.AxisColor = Color.Brown;
                    rule.GradientFill = true;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)); // C1:C11
                    rule = new XlCondFmtRuleDataBar();
                    rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.2);
                    rule.MinLength = 10;
                    rule.MaxLength = 90;
                    rule.MinValue.ObjectType = XlCondFmtValueObjectType.Number;
                    rule.MinValue.Value = 3;
                    rule.Direction = XlDataBarDirection.RightToLeft;
                    rule.ShowValues = false;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #DataBarRule
        }

        static void IconSet(Stream stream, XlDocumentFormat documentFormat) {
            #region #IconSetRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    if(cell.ColumnIndex % 2 == 0)
                                        cell.Value = cell.RowIndex + 1;
                                    else
                                        cell.Value = cell.RowIndex - 5;
                                }
                            }
                        }
                    }

                    // Format values using icon set rule
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); // A1:A11
                    XlCondFmtRuleIconSet rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Arrows3;
                    rule.Priority = 1;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)); // B1:B11
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Flags3;
                    rule.Priority = 2;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)); // C1:C11
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.Rating5;
                    rule.ShowValues = false;
                    rule.Priority = 3;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)); // D1:D11
                    rule = new XlCondFmtRuleIconSet();
                    rule.IconSetType = XlCondFmtIconSetType.TrafficLights4;
                    rule.Reverse = true;
                    rule.Priority = 4;
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #IconSetRule
        }

        static void ColorScale(Stream stream, XlDocumentFormat documentFormat) {
            #region #ColorScaleRule
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create sample content
                    for(int i = 0; i < 11; i++) {
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                using(IXlCell cell = row.CreateCell()) {
                                    cell.Value = cell.RowIndex + 1;
                                }
                            }
                        }
                    }

                    // Format values using color scale rule
                    XlConditionalFormatting formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)); // A1:A11
                    formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)); // C1:C11
                    XlCondFmtRuleColorScale rule = new XlCondFmtRuleColorScale();
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);

                    formatting = new XlConditionalFormatting();
                    formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)); // B1:B11
                    formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)); // D1:D11
                    rule = new XlCondFmtRuleColorScale();
                    rule.ColorScaleType = XlCondFmtColorScaleType.ColorScale2;
                    rule.MinColor = XlColor.FromTheme(XlThemeColor.Light1, 0.0);
                    rule.MaxColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.5);
                    formatting.Rules.Add(rule);
                    sheet.ConditionalFormattings.Add(formatting);
                }
            }

            #endregion #ColorScaleRule
        }

    }
}