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
	Public NotInheritable Class Miscellaneous

		Private Sub New()
		End Sub
		Private Shared Sub Hyperlinks(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Hyperlinks"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 300
					End Using

					' Local link
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Local link"
							cell.Formatting = XlCellFormatting.Hyperlink
							Dim hyperlink As New XlHyperlink()
							hyperlink.Reference = New XlCellRange(New XlCellPosition(cell.ColumnIndex, cell.RowIndex))
							hyperlink.TargetUri = "#Sheet1!C5"
							sheet.Hyperlinks.Add(hyperlink)
						End Using
					End Using

					' External file link
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "External file link"
							cell.Formatting = XlCellFormatting.Hyperlink
							Dim hyperlink As New XlHyperlink()
							hyperlink.Reference = New XlCellRange(New XlCellPosition(cell.ColumnIndex, cell.RowIndex))
							hyperlink.TargetUri = "linked.xlsx#Sheet1!C5"
							sheet.Hyperlinks.Add(hyperlink)
						End Using
					End Using

					' External URI
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "External uri"
							cell.Formatting = XlCellFormatting.Hyperlink
							Dim hyperlink As New XlHyperlink()
							hyperlink.Reference = New XlCellRange(New XlCellPosition(cell.ColumnIndex, cell.RowIndex))
							hyperlink.TargetUri = "http://www.devexpress.com"
							sheet.Hyperlinks.Add(hyperlink)
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #Hyperlinks
		End Sub

		Private Shared Sub DocumentProperties(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#DocumentProperties"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Setup built-in document properties
				document.Properties.Title = "Sample document"
				document.Properties.Subject = "XL Export API demo"
				document.Properties.Keywords = "XL export document generation"
				document.Properties.Description = "Generate through XL Export API"
				document.Properties.Category = "Spreadsheet"
				document.Properties.Company = "DevExpress Inc."

				' Setup custom properties
				document.Properties.Custom("Product Suite") = "Spreadsheet Document Automation"
				document.Properties.Custom("Revision") = 5
				document.Properties.Custom("Date Completed") = DateTime.Now
				document.Properties.Custom("Published") = True

				' Generate sample content
				Using sheet As IXlSheet = document.CreateSheet()
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						row.SkipCells(1)
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "You can check exported document properties using File/Info/Advanced Properties dialog box."
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #DocumentProperties
		End Sub

		Private Shared Sub DocumentOptions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#DocumentOptions"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create sheet
				Using sheet As IXlSheet = document.CreateSheet()
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 200
					End Using
					Using column As IXlColumn = sheet.CreateColumn()
						column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Bottom)
					End Using
					' Document format
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Document format:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.DocumentFormat.ToString().ToUpper()
						End Using
					End Using
					' Max column count
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Max column count:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.MaxColumnCount
						End Using
					End Using
					' Max row count
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Max row count:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.MaxRowCount
						End Using
					End Using
					' Supports document parts
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Supports document parts:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.SupportsDocumentParts
						End Using
					End Using
					' Supports formulas
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Supports formulas:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.SupportsFormulas
						End Using
					End Using
					' Supports outline/grouping
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Supports outline/grouping:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = document.Options.SupportsOutlineGrouping
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #DocumentOptions
		End Sub

		Private Shared Sub CsvExportOptions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CsvOptions"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Setup CSV export options
				Dim csvOptions As CsvDataAwareExporterOptions = TryCast(document.Options, CsvDataAwareExporterOptions)
				If csvOptions IsNot Nothing Then
					csvOptions.Encoding = Encoding.UTF8
					csvOptions.WritePreamble = True
					csvOptions.UseCellNumberFormat = False
					csvOptions.NewlineAfterLastRow = True
				End If

				' Create sample content
				Using sheet As IXlSheet = document.CreateSheet()
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
							End Using
						Next i
					End Using
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #CsvOptions
		End Sub

	End Class
End Namespace