Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
	Public NotInheritable Class BasicActions

		Private Sub New()
		End Sub
		Private Shared Sub CreateDocument(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateDocument"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture
			End Using

'			#End Region ' #CreateDocument
		End Sub

		Private Shared Sub CreateSheet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateSheet"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Sales report"
				End Using
			End Using

'			#End Region ' #CreateSheet
		End Sub

		Private Shared Sub CreateHiddenSheet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateHiddenSheet"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Sales report"
				End Using

				' Create worksheet and setup visibility
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.Name = "Data"
					sheet.VisibleState = XlSheetVisibleState.Hidden
				End Using
			End Using

'			#End Region ' #CreateHiddenSheet
		End Sub

		Private Shared Sub CreateColumns(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateColumns"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Set column A width to 100 pixels
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 100
					End Using

					' Hide column B
					Using column As IXlColumn = sheet.CreateColumn()
						column.IsHidden = True
					End Using

					' Set column D width of 24.5 characters
					Using column As IXlColumn = sheet.CreateColumn(3)
						column.WidthInCharacters = 24.5f
					End Using
				End Using
			End Using

'			#End Region ' #CreateColumns
		End Sub

		Private Shared Sub CreateRows(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateRows"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Set row 1 height to 40 pixels
					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 40
					End Using

					' Hide row 3
					Using row As IXlRow = sheet.CreateRow(2)
						row.IsHidden = True
					End Using
				End Using
			End Using

'			#End Region ' #CreateRows
		End Sub

		Private Shared Sub CreateCells(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CreateCells"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)

				' Setup document culture
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 150
					End Using

					Using row As IXlRow = sheet.CreateRow()

						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Numeric value:"
						End Using

						' Create cell with numeric value
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 123.45
						End Using
					End Using

					Using row As IXlRow = sheet.CreateRow()

						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Text value:"
						End Using

						' Create cell with text value
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "abc"
						End Using
					End Using

					Using row As IXlRow = sheet.CreateRow()

						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Boolean value:"
						End Using

						' Create cell with boolean value
						Using cell As IXlCell = row.CreateCell()
							cell.Value = True
						End Using
					End Using

					Using row As IXlRow = sheet.CreateRow()

						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Error value:"
						End Using

						' Create cell with error value
						Using cell As IXlCell = row.CreateCell()
							cell.Value = XlVariantValue.ErrorName
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #CreateCells
		End Sub

		Private Shared Sub MergeCells(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#MergeCells"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				Using sheet As IXlSheet = document.CreateSheet()
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Merged cells in range A1:E1"
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
						End Using
					End Using
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Merged cells in range A2:A5"
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
							cell.Formatting.Alignment.WrapText = True
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Merged cells in range B2:E5"
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
						End Using
					End Using

					' Merge cells in range A1:E1
					sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 0, 4, 0))

					' Merge cells in range A2:A5
					sheet.MergedCells.Add(XlCellRange.FromLTRB(0, 1, 0, 4))

					' Merge cells in range B2:E5
					sheet.MergedCells.Add(XlCellRange.FromLTRB(1, 1, 4, 4))
				End Using
			End Using

'			#End Region ' #MergeCells
		End Sub

	End Class
End Namespace