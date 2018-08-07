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
    public static class Pictures {

        static void InsertPicture(Stream stream, XlDocumentFormat documentFormat, string imagesPath) {
            #region #InsertPicture
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {

                    // Load picture from file and insert it with two cell anchor
                    using(IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.SetTwoCellAnchor(new XlAnchorPoint(1, 1, 0, 0), new XlAnchorPoint(6, 11, 2, 15), XlAnchorType.TwoCell);
                    }
                }
            }

            #endregion #InsertPicture
        }

        static void StretchPicture(Stream stream, XlDocumentFormat documentFormat, string imagesPath) {
            #region #StretchPicture
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.SkipColumns(1);
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 205;
                    }
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 154;
                    }

                    // Load picture from file and stretch it into specified cell
                    using(IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.StretchToCell(new XlCellPosition(1, 1)); // B2
                    }
                }
            }

            #endregion #StretchPicture
        }

        static void FitPicture(Stream stream, XlDocumentFormat documentFormat, string imagesPath) {
            #region #FitPicture
            // Create exporter
            IXlExporter exporter = XlExport.CreateExporter(documentFormat);

            // Create document
            using(IXlDocument document = exporter.CreateDocument(stream)) {
                document.Options.Culture = CultureInfo.CurrentCulture;

                // Create worksheet
                using(IXlSheet sheet = document.CreateSheet()) {
                    sheet.SkipColumns(1);
                    using(IXlColumn column = sheet.CreateColumn()) {
                        column.WidthInPixels = 300;
                    }
                    sheet.SkipRows(1);
                    using(IXlRow row = sheet.CreateRow()) {
                        row.HeightInPixels = 154;
                    }

                    // Load picture from file and fit it into specified cell
                    using(IXlPicture picture = sheet.CreatePicture()) {
                        picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"));
                        picture.FitToCell(new XlCellPosition(1, 1), 300, 154, true);
                    }
                }
            }

            #endregion #FitPicture
        }

    }
}