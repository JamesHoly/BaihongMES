Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
	Public NotInheritable Class PageViewAndLayout

		Private Sub New()
		End Sub
		Private Shared Sub FreezeRow(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#FreezeRow"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Freeze first row
					sheet.SplitPosition = New XlCellPosition(0, 1)

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					For i As Integer = 0 To 3
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
									cell.ApplyFormatting(rowFormatting)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #FreezeRow
		End Sub

		Private Shared Sub FreezeColumn(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#FreezeColumn"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Freeze first column
					sheet.SplitPosition = New XlCellPosition(1, 0)

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					For i As Integer = 0 To 3
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
									cell.ApplyFormatting(rowFormatting)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #FreezeColumn
		End Sub

		Private Shared Sub FreezePanes(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#FreezePanes"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Freeze first column and first row
					sheet.SplitPosition = New XlCellPosition(1, 1)

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					For i As Integer = 0 To 3
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
									cell.ApplyFormatting(rowFormatting)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #FreezePanes
		End Sub

		Private Shared Sub HeadersFooters(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#HeadersAndFooters"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Setup headers/footers
					sheet.HeaderFooter.DifferentOddEven = True
					sheet.HeaderFooter.OddHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.Bold & "Sample report", Nothing, XlHeaderFooter.BookName)
					sheet.HeaderFooter.OddFooter = XlHeaderFooter.FromLCR(Nothing, Nothing, XlHeaderFooter.PageNumber)
					sheet.HeaderFooter.EvenHeader = XlHeaderFooter.FromLCR(XlHeaderFooter.BookPath, Nothing, XlHeaderFooter.SheetName)
					sheet.HeaderFooter.EvenFooter = XlHeaderFooter.FromLCR(XlHeaderFooter.PageNumber, Nothing, XlHeaderFooter.Date)

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					For i As Integer = 0 To 3
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
									cell.ApplyFormatting(rowFormatting)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #HeaderAndFooters
		End Sub

		Private Shared Sub PageBreaks(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PageBreaks"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Setup manual page breaks
					sheet.ColumnPageBreaks.Add(2) ' Column C
					sheet.RowPageBreaks.Add(10) ' Row 11

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 100
						column.Formatting = New XlCellFormatting()
						column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
					End Using

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Sales"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
								cell.ApplyFormatting(rowFormatting)
							End Using
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #PageBreaks
		End Sub

		Private Shared Sub PageMargins(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PageMargins"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Setup page margins
					sheet.PageMargins = New XlPageMargins()
					sheet.PageMargins.PageUnits = XlPageUnits.Centimeters
					sheet.PageMargins.Left = 2.0
					sheet.PageMargins.Right = 1.0
					sheet.PageMargins.Top = 1.25
					sheet.PageMargins.Bottom = 1.25
					sheet.PageMargins.Header = 0.7
					sheet.PageMargins.Footer = 0.7

					' Generate sample content
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						row.SkipCells(1)
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "You can check page margins using Page Setup dialog box."
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #PageMargins
		End Sub

		Private Shared Sub PageSetup(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PageSetup"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Specify page setup settings
					sheet.PageSetup = New XlPageSetup()
					sheet.PageSetup.PaperKind = System.Drawing.Printing.PaperKind.A4
					sheet.PageSetup.PageOrientation = XlPageOrientation.Landscape
					sheet.PageSetup.FitToPage = True
					sheet.PageSetup.FitToWidth = 1
					sheet.PageSetup.FitToHeight = 0
					sheet.PageSetup.BlackAndWhite = True
					sheet.PageSetup.Copies = 2

					' Generate sample content
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						row.SkipCells(1)
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "You can check settings using Page Setup dialog box."
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #PageSetup
		End Sub

		Private Shared Sub PrintArea(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PrintArea"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Set print area
					sheet.PrintArea = XlCellRange.FromLTRB(0, 0, 4, 4) ' A1:E5

					' Create columns
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 110
						column.Formatting = XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom)
					End Using
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 190
					End Using
					For i As Integer = 0 To 1
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 90
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 130
					End Using
					sheet.SkipColumns(1)
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 130
					End Using

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True
					headerRowFormatting.Font.Color = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
					headerRowFormatting.Fill = XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent2, 0.0))

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Employee ID"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Employee name"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Salary"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Bonus"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Department"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						row.SkipCells(1)
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Departments"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
					End Using

					' Generate sample content
					Dim id() As Integer = { 10115, 10709, 10401, 10204 }
					Dim name() As String = { "Augusta Delono", "Chris Cadwell", "Frank Diamond", "Simon Newman" }
					Dim salary() As Integer = { 1100, 2000, 1750, 1250 }
					Dim bonus() As Integer = { 50, 180, 100, 80 }
					Dim deptid() As Integer = { 0, 2, 3, 3 }
					Dim department() As String = { "Accounting", "IT", "Management", "Manufacturing" }
					For i As Integer = 0 To 3
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = id(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = name(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = salary(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = bonus(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = department(deptid(i))
								cell.ApplyFormatting(rowFormatting)
							End Using
							row.SkipCells(1)
							Using cell As IXlCell = row.CreateCell()
								cell.Value = department(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
						End Using
					Next i

					Dim validation As New XlDataValidation()
					validation.Ranges.Add(XlCellRange.FromLTRB(4, 1, 4, 4)) ' E2:E5
					validation.Type = XlDataValidationType.List
					validation.Criteria1 = XlCellRange.FromLTRB(6, 1, 6, 4).AsAbsolute() ' $G$2:$G$5
					sheet.DataValidations.Add(validation)
				End Using
			End Using

'			#End Region ' #PrintArea
		End Sub

		Private Shared Sub PrintOptions(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PrintOptions"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Specify print options
					sheet.PrintOptions = New XlPrintOptions()
					sheet.PrintOptions.Headings = True
					sheet.PrintOptions.GridLines = True
					sheet.PrintOptions.HorizontalCentered = True
					sheet.PrintOptions.VerticalCentered = True

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 100
						column.Formatting = New XlCellFormatting()
						column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
					End Using

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Sales"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
								cell.ApplyFormatting(rowFormatting)
							End Using
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #PrintOptions
		End Sub

		Private Shared Sub PrintTitles(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PrintTitles"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Set first column and first row as print titles
					sheet.PrintTitles.SetRows(0, 0)
					sheet.PrintTitles.SetColumns(0, 0)

					' Generate sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 250
					End Using
					For i As Integer = 0 To 3
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
							column.Formatting = New XlCellFormatting()
							column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
					Next i

					' Prepare cells formatting
					Dim rowFormatting As New XlCellFormatting()
					rowFormatting.Font = New XlFont()
					rowFormatting.Font.Name = "Century Gothic"
					rowFormatting.Font.SchemeStyle = XlFontSchemeStyles.None

					' Prepare header row formatting
					Dim headerRowFormatting As New XlCellFormatting()
					headerRowFormatting.CopyFrom(rowFormatting)
					headerRowFormatting.Font.Bold = True

					' Generate header row
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Product"
							cell.ApplyFormatting(headerRowFormatting)
						End Using
						For i As Integer = 0 To 3
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Q{0}", i + 1)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using

					' Generate data rows
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Gnocchi di nonna Alice", "Gudbrandsdalsost", "Gustaf's Knackebrod", "Queso Cabrales", "Queso Manchego La Pastora", "Raclette Courdavault", "Singaporean Hokkien Fried Mee", "Wimmers gute Semmelknodel" }
					Dim random As New Random()
					For i As Integer = 0 To 11
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
								cell.ApplyFormatting(rowFormatting)
							End Using
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = Math.Round(random.NextDouble() * 2000 + 3000)
									cell.ApplyFormatting(rowFormatting)
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #PrintTitles
		End Sub

	End Class
End Namespace