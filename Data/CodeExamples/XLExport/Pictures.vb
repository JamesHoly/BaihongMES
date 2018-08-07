Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports DevExpress.Export.Xl
Imports DevExpress.XtraExport.Csv
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
	Public NotInheritable Class Pictures

		Private Sub New()
		End Sub
		Private Shared Sub InsertPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat, ByVal imagesPath As String)
'			#Region "#InsertPicture"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Load picture from file and insert it with two cell anchor
					Using picture As IXlPicture = sheet.CreatePicture()
						picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
						picture.SetTwoCellAnchor(New XlAnchorPoint(1, 1, 0, 0), New XlAnchorPoint(6, 11, 2, 15), XlAnchorType.TwoCell)
					End Using
				End Using
			End Using

'			#End Region ' #InsertPicture
		End Sub

		Private Shared Sub StretchPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat, ByVal imagesPath As String)
'			#Region "#StretchPicture"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.SkipColumns(1)
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 205
					End Using
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 154
					End Using

					' Load picture from file and stretch it into specified cell
					Using picture As IXlPicture = sheet.CreatePicture()
						picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
						picture.StretchToCell(New XlCellPosition(1, 1)) ' B2
					End Using
				End Using
			End Using

'			#End Region ' #StretchPicture
		End Sub

		Private Shared Sub FitPicture(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat, ByVal imagesPath As String)
'			#Region "#FitPicture"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.SkipColumns(1)
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 300
					End Using
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 154
					End Using

					' Load picture from file and fit it into specified cell
					Using picture As IXlPicture = sheet.CreatePicture()
						picture.Image = Image.FromFile(Path.Combine(imagesPath, "image1.jpg"))
						picture.FitToCell(New XlCellPosition(1, 1), 300, 154, True)
					End Using
				End Using
			End Using

'			#End Region ' #FitPicture
		End Sub

	End Class
End Namespace