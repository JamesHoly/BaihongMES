using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using DevExpress.Export.Xl;
using DevExpress.XtraExport.Csv;
using DevExpress.Spreadsheet;

namespace XLExportExamples {
    public static class CellFormatting {

        static void PredefinedFormatting(Stream stream, XlDocumentFormat documentFormat) {
            #region #PredefinedFormatting
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);
            
            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                using(IXlSheet sheet = document.CreateSheet()) {
                    
                    // Create columns
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }
                    
                    // Good, Bad, Neutral
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Good, Bad, Neutral";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Normal";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Bad";
                            cell.Formatting = XlCellFormatting.Bad;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Good";
                            cell.Formatting = XlCellFormatting.Good;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Neutral";
                            cell.Formatting = XlCellFormatting.Neutral;
                        }
                    }

                    // Data and Model
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Data and Model";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Calculation";
                            cell.Formatting = XlCellFormatting.Calculation;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Check Cell";
                            cell.Formatting = XlCellFormatting.CheckCell;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Explanatory";
                            cell.Formatting = XlCellFormatting.Explanatory;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Input";
                            cell.Formatting = XlCellFormatting.Input;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Linked Cell";
                            cell.Formatting = XlCellFormatting.LinkedCell;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Note";
                            cell.Formatting = XlCellFormatting.Note;
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Output";
                            cell.Formatting = XlCellFormatting.Output;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Warning Text";
                            cell.Formatting = XlCellFormatting.WarningText;
                        }
                    }

                    // Titles and Headings
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Titles and Headings";
                        }
                    }
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading1";
                            cell.Formatting = XlCellFormatting.Heading1;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading2";
                            cell.Formatting = XlCellFormatting.Heading2;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading3";
                            cell.Formatting = XlCellFormatting.Heading3;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Heading4";
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Title";
                            cell.Formatting = XlCellFormatting.Title;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Total";
                            cell.Formatting = XlCellFormatting.Total;
                        }
                    }
                }
            }

            #endregion #PredefinedFormatting
        }

        static void ThemedFormatting(Stream stream, XlDocumentFormat documentFormat) {
            #region #ThemedFormatting
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create columns
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }

                    XlThemeColor[] themeColors = new XlThemeColor[] { XlThemeColor.Accent1, XlThemeColor.Accent2, XlThemeColor.Accent3, XlThemeColor.Accent4, XlThemeColor.Accent5, XlThemeColor.Accent6 };

                    // 20% of theme accent colors
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Accent{0} 20%", i + 1);
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.8);
                            }
                        }
                    }

                    // 40% of theme accent colors
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Accent{0} 40%", i + 1);
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.6);
                            }
                        }
                    }

                    // 60% of theme accent colors
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Accent{0} 60%", i + 1);
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.4);
                            }
                        }
                    }

                    // Theme accent colors
                    using(IXlRow row = sheet.CreateRow()) {
                        for(int i = 0; i < 6; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("Accent{0}", i + 1);
                                cell.Formatting = XlCellFormatting.Themed(themeColors[i], 0.0);
                            }
                        }
                    }
                }
            }

            #endregion #ThemedFormatting
        }

        static void Alignment(Stream stream, XlDocumentFormat documentFormat) {
            #region #Alignment
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Create columns
                    for(int i = 0; i < 3; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 72;
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 40;
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Left/Top alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Top));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Center/Top alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Top));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Right/Top alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Top));
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 40;
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Left/Center alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Center/Center alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Right/Center alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center));
                        }
                    }

                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 40;
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Left/Bottom alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Center/Bottom alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Bottom));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "text";
                            // Right/Bottom alignment
                            cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom));
                        }
                    }

                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Wrapped text";
                            // Wrapped text
                            cell.Formatting = new XlCellAlignment() { WrapText = true };
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Indented text";
                            // Indented text
                            cell.Formatting = new XlCellAlignment() { Indent = 2 };
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Rotated text";
                            // Rotated text
                            cell.Formatting = new XlCellAlignment() { TextRotation = 90 };
                        }
                    }
                }
            }

            #endregion #Alignment
        }

        static void Borders(Stream stream, XlDocumentFormat documentFormat) {
            #region #Borders
            XlBorderLineStyle[,] lineStyles = new XlBorderLineStyle[,] {
                        { XlBorderLineStyle.Thin, XlBorderLineStyle.Medium, XlBorderLineStyle.Thick, XlBorderLineStyle.Double },
                        { XlBorderLineStyle.Dotted, XlBorderLineStyle.Dashed, XlBorderLineStyle.DashDot, XlBorderLineStyle.DashDotDot },
                        { XlBorderLineStyle.SlantDashDot, XlBorderLineStyle.MediumDashed, XlBorderLineStyle.MediumDashDot, XlBorderLineStyle.MediumDashDotDot }
                    };

            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    for(int i = 0; i < 3; i++) {
                        sheet.SkipRows(1);
                        using(IXlRow row = sheet.CreateRow()) {
                            for(int j = 0; j < 4; j++) {
                                row.SkipCells(1);
                                using(IXlCell cell = row.CreateCell()) {
                                    // Set cell borders
                                    cell.ApplyFormatting(XlBorder.OutlineBorders(Color.Black, lineStyles[i, j]));
                                }
                            }
                        }
                    }
                }
            }

            #endregion #Borders
        }

        static void Fill(Stream stream, XlDocumentFormat documentFormat) {
            #region #Fill
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Set solid fill of known color
                            cell.ApplyFormatting(XlFill.SolidFill(Color.Beige));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set solid fill of RGB color
                            cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(0xff, 0x99, 0x66)));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set solid fill of themed color
                            cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent3, 0.4)));
                        }
                    }

                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            // Set pattern fill of known colors
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkDown, Color.Red, Color.White));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set pattern fill of RGB colors
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkTrellis, Color.FromArgb(0xff, 0xff, 0x66), Color.FromArgb(0x66, 0x99, 0xff)));
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            // Set pattern fill of themed colors
                            cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.LightHorizontal, XlColor.FromTheme(XlThemeColor.Accent1, 0.2), XlColor.FromTheme(XlThemeColor.Light2, 0.0)));
                        }
                    }
                }
            }

            #endregion #Fill
        }

        static void Font(Stream stream, XlDocumentFormat documentFormat) {
            #region #Font
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    for(int i = 0; i < 5; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 100;
                        }
                    }

                    // Set font name / font scheme
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Body font";
                            // Set body font
                            cell.ApplyFormatting(XlFont.BodyFont());
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Headings font";
                            // Set headings font
                            cell.ApplyFormatting(XlFont.HeadingsFont());
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Custom font";
                            // Set custom font
                            XlFont font = new XlFont();
                            font.Name = "Century Gothic";
                            font.SchemeStyle = XlFontSchemeStyles.None;
                            cell.ApplyFormatting(font);
                        }
                    }

                    // Set font size
                    int[] fontSizes = new int[] { 11, 14, 18, 24, 36 };
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {

                        for(int i = 0; i < 5; i++) {
                            using(IXlCell cell = row.CreateCell()) {
                                cell.Value = string.Format("{0}pt", fontSizes[i]);
                                XlFont font = new XlFont();
                                font.Size = fontSizes[i];
                                cell.ApplyFormatting(font);
                            }
                        }
                    }

                    // Set font parameters
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        // Color
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Red";
                            XlFont font = new XlFont() { Color = Color.Red };
                            cell.ApplyFormatting(font);
                        }
                        // Bold
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Bold";
                            XlFont font = new XlFont() { Bold = true };
                            cell.ApplyFormatting(font);
                        }
                        // Italic
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Italic";
                            XlFont font = new XlFont() { Italic = true };
                            cell.ApplyFormatting(font);
                        }
                        // Underline
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Underline";
                            XlFont font = new XlFont() { Underline = XlUnderlineType.Double };
                            cell.ApplyFormatting(font);
                        }
                        // Strike through
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "StrikeThrough";
                            XlFont font = new XlFont() { StrikeThrough = true };
                            cell.ApplyFormatting(font);
                        }
                    }
                }
            }

            #endregion #Font
        }

        static void NumberFormat(Stream stream, XlDocumentFormat documentFormat) {
            #region #NumberFormat
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;
                CsvDataAwareExporterOptions csvOptions = document.Options as CsvDataAwareExporterOptions;
                if(csvOptions != null)
                    csvOptions.WritePreamble = true;
                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    for(int i = 0; i < 6; i++) {
                        using(IXlColumn column = sheet.CreateColumn()) {
                            column.WidthInPixels = 180;
                        }
                    }

                    // Excel number formats
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Excel number formats";
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                    }
                    // Predefined formats
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Predefined formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 123.45;
                            cell.Formatting = XlNumberFormat.Number2;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 12345;
                            cell.Formatting = XlNumberFormat.NumberWithThousandSeparator;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 0.33;
                            cell.Formatting = XlNumberFormat.Percentage;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlNumberFormat.ShortDate;
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlNumberFormat.ShortTime12;
                        }
                    }
                    // Custom formats
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Custom formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 4310.45;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = @"_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 3426.75;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = @"_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "" - ""??_-;_-@_-";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 0.333;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "0.0%";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "dddd, mmmm d, yyyy";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 0.6234;
                            cell.Formatting = new XlCellFormatting();
                            cell.Formatting.NumberFormat = "# ???/???";
                        }
                    }

                    // .NET number formats
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = ".NET number formats";
                            cell.Formatting = XlCellFormatting.Heading4;
                        }
                    }
                    // Standard format strings
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Standard formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 123.45;
                            cell.Formatting = XlCellFormatting.FromNetFormat("D", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 12345;
                            cell.Formatting = XlCellFormatting.FromNetFormat("E", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 0.33;
                            cell.Formatting = XlCellFormatting.FromNetFormat("P", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("d", true);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("t", true);
                        }
                    }
                    // Custom format strings
                    using(IXlRow row = sheet.CreateRow()) {
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = "Custom formats:";
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 123.45;
                            cell.Formatting = XlCellFormatting.FromNetFormat("#0.00", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 12345;
                            cell.Formatting = XlCellFormatting.FromNetFormat("0.0##e+00", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = 0.333;
                            cell.Formatting = XlCellFormatting.FromNetFormat("Max={0:#.0%}", false);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("dd-MM-yyyy", true);
                        }
                        using(IXlCell cell = row.CreateCell()) {
                            cell.Value = DateTime.Now;
                            cell.Formatting = XlCellFormatting.FromNetFormat("hh:mm tt", true);
                        }
                    }
                }
            }

            #endregion #NumberFormat
        }
    }
}