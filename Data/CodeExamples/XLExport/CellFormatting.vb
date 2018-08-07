Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.XtraExport.Csv
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
	Public NotInheritable Class CellFormatting

		Private Sub New()
		End Sub
		Private Shared Sub PredefinedFormatting(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#PredefinedFormatting"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				Using sheet As IXlSheet = document.CreateSheet()

					' Create columns
					For i As Integer = 0 To 5
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
						End Using
					Next i

					' Good, Bad, Neutral
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Good, Bad, Neutral"
						End Using
					End Using
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Normal"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Bad"
							cell.Formatting = XlCellFormatting.Bad
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Good"
							cell.Formatting = XlCellFormatting.Good
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Neutral"
							cell.Formatting = XlCellFormatting.Neutral
						End Using
					End Using

					' Data and Model
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Data and Model"
						End Using
					End Using
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Calculation"
							cell.Formatting = XlCellFormatting.Calculation
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Check Cell"
							cell.Formatting = XlCellFormatting.CheckCell
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Explanatory"
							cell.Formatting = XlCellFormatting.Explanatory
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Input"
							cell.Formatting = XlCellFormatting.Input
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Linked Cell"
							cell.Formatting = XlCellFormatting.LinkedCell
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Note"
							cell.Formatting = XlCellFormatting.Note
						End Using
					End Using
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Output"
							cell.Formatting = XlCellFormatting.Output
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Warning Text"
							cell.Formatting = XlCellFormatting.WarningText
						End Using
					End Using

					' Titles and Headings
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Titles and Headings"
						End Using
					End Using
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Heading1"
							cell.Formatting = XlCellFormatting.Heading1
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Heading2"
							cell.Formatting = XlCellFormatting.Heading2
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Heading3"
							cell.Formatting = XlCellFormatting.Heading3
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Heading4"
							cell.Formatting = XlCellFormatting.Heading4
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Title"
							cell.Formatting = XlCellFormatting.Title
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Total"
							cell.Formatting = XlCellFormatting.Total
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #PredefinedFormatting
		End Sub

		Private Shared Sub ThemedFormatting(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#ThemedFormatting"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				Using sheet As IXlSheet = document.CreateSheet()

					' Create columns
					For i As Integer = 0 To 5
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
						End Using
					Next i

					Dim themeColors() As XlThemeColor = { XlThemeColor.Accent1, XlThemeColor.Accent2, XlThemeColor.Accent3, XlThemeColor.Accent4, XlThemeColor.Accent5, XlThemeColor.Accent6 }

					' 20% of theme accent colors
					Using row As IXlRow = sheet.CreateRow()
						For i As Integer = 0 To 5
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Accent{0} 20%", i + 1)
								cell.Formatting = XlCellFormatting.Themed(themeColors(i), 0.8)
							End Using
						Next i
					End Using

					' 40% of theme accent colors
					Using row As IXlRow = sheet.CreateRow()
						For i As Integer = 0 To 5
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Accent{0} 40%", i + 1)
								cell.Formatting = XlCellFormatting.Themed(themeColors(i), 0.6)
							End Using
						Next i
					End Using

					' 60% of theme accent colors
					Using row As IXlRow = sheet.CreateRow()
						For i As Integer = 0 To 5
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Accent{0} 60%", i + 1)
								cell.Formatting = XlCellFormatting.Themed(themeColors(i), 0.4)
							End Using
						Next i
					End Using

					' Theme accent colors
					Using row As IXlRow = sheet.CreateRow()
						For i As Integer = 0 To 5
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("Accent{0}", i + 1)
								cell.Formatting = XlCellFormatting.Themed(themeColors(i), 0.0)
							End Using
						Next i
					End Using
				End Using
			End Using

'			#End Region ' #ThemedFormatting
		End Sub

		Private Shared Sub Alignment(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Alignment"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create columns
					For i As Integer = 0 To 2
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 72
						End Using
					Next i

					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 40
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Left/Top alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Top))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Center/Top alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Top))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Right/Top alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Top))
						End Using
					End Using

					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 40
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Left/Center alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Center))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Center/Center alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Center))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Right/Center alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Center))
						End Using
					End Using

					Using row As IXlRow = sheet.CreateRow()
						row.HeightInPixels = 40
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Left/Bottom alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Left, XlVerticalAlignment.Bottom))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Center/Bottom alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Center, XlVerticalAlignment.Bottom))
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "text"
							' Right/Bottom alignment
							cell.ApplyFormatting(XlCellAlignment.FromHV(XlHorizontalAlignment.Right, XlVerticalAlignment.Bottom))
						End Using
					End Using

					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Wrapped text"
							' Wrapped text
							cell.Formatting = New XlCellAlignment() With {.WrapText = True}
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Indented text"
							' Indented text
							cell.Formatting = New XlCellAlignment() With {.Indent = 2}
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Rotated text"
							' Rotated text
							cell.Formatting = New XlCellAlignment() With {.TextRotation = 90}
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #Alignment
		End Sub

		Private Shared Sub Borders(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Borders"
			Dim lineStyles(,) As XlBorderLineStyle = { { XlBorderLineStyle.Thin, XlBorderLineStyle.Medium, XlBorderLineStyle.Thick, XlBorderLineStyle.Double }, { XlBorderLineStyle.Dotted, XlBorderLineStyle.Dashed, XlBorderLineStyle.DashDot, XlBorderLineStyle.DashDotDot }, { XlBorderLineStyle.SlantDashDot, XlBorderLineStyle.MediumDashed, XlBorderLineStyle.MediumDashDot, XlBorderLineStyle.MediumDashDotDot } }

			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					For i As Integer = 0 To 2
						sheet.SkipRows(1)
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								row.SkipCells(1)
								Using cell As IXlCell = row.CreateCell()
									' Set cell borders
									cell.ApplyFormatting(XlBorder.OutlineBorders(Color.Black, lineStyles(i, j)))
								End Using
							Next j
						End Using
					Next i
				End Using
			End Using

'			#End Region ' #Borders
		End Sub

		Private Shared Sub Fill(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Fill"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							' Set solid fill of known color
							cell.ApplyFormatting(XlFill.SolidFill(Color.Beige))
						End Using
						Using cell As IXlCell = row.CreateCell()
							' Set solid fill of RGB color
							cell.ApplyFormatting(XlFill.SolidFill(Color.FromArgb(&Hff, &H99, &H66)))
						End Using
						Using cell As IXlCell = row.CreateCell()
							' Set solid fill of themed color
							cell.ApplyFormatting(XlFill.SolidFill(XlColor.FromTheme(XlThemeColor.Accent3, 0.4)))
						End Using
					End Using

					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							' Set pattern fill of known colors
							cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkDown, Color.Red, Color.White))
						End Using
						Using cell As IXlCell = row.CreateCell()
							' Set pattern fill of RGB colors
							cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.DarkTrellis, Color.FromArgb(&Hff, &Hff, &H66), Color.FromArgb(&H66, &H99, &Hff)))
						End Using
						Using cell As IXlCell = row.CreateCell()
							' Set pattern fill of themed colors
							cell.ApplyFormatting(XlFill.PatternFill(XlPatternType.LightHorizontal, XlColor.FromTheme(XlThemeColor.Accent1, 0.2), XlColor.FromTheme(XlThemeColor.Light2, 0.0)))
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #Fill
		End Sub

		Private Shared Sub Font(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Font"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					For i As Integer = 0 To 4
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
						End Using
					Next i

					' Set font name / font scheme
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Body font"
							' Set body font
							cell.ApplyFormatting(XlFont.BodyFont())
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Headings font"
							' Set headings font
							cell.ApplyFormatting(XlFont.HeadingsFont())
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Custom font"
							' Set custom font
							Dim font As New XlFont()
							font.Name = "Century Gothic"
							font.SchemeStyle = XlFontSchemeStyles.None
							cell.ApplyFormatting(font)
						End Using
					End Using

					' Set font size
					Dim fontSizes() As Integer = { 11, 14, 18, 24, 36 }
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()

						For i As Integer = 0 To 4
							Using cell As IXlCell = row.CreateCell()
								cell.Value = String.Format("{0}pt", fontSizes(i))
								Dim font As New XlFont()
								font.Size = fontSizes(i)
								cell.ApplyFormatting(font)
							End Using
						Next i
					End Using

					' Set font parameters
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						' Color
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Red"
							Dim font As New XlFont() With {.Color = Color.Red}
							cell.ApplyFormatting(font)
						End Using
						' Bold
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Bold"
							Dim font As New XlFont() With {.Bold = True}
							cell.ApplyFormatting(font)
						End Using
						' Italic
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Italic"
							Dim font As New XlFont() With {.Italic = True}
							cell.ApplyFormatting(font)
						End Using
						' Underline
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Underline"
							Dim font As New XlFont() With {.Underline = XlUnderlineType.Double}
							cell.ApplyFormatting(font)
						End Using
						' Strike through
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "StrikeThrough"
							Dim font As New XlFont() With {.StrikeThrough = True}
							cell.ApplyFormatting(font)
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #Font
		End Sub

		Private Shared Sub NumberFormat(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#NumberFormat"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture
				Dim csvOptions As CsvDataAwareExporterOptions = TryCast(document.Options, CsvDataAwareExporterOptions)
				If csvOptions IsNot Nothing Then
					csvOptions.WritePreamble = True
				End If
				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()
					For i As Integer = 0 To 5
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 180
						End Using
					Next i

					' Excel number formats
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Excel number formats"
							cell.Formatting = XlCellFormatting.Heading4
						End Using
					End Using
					' Predefined formats
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Predefined formats:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 123.45
							cell.Formatting = XlNumberFormat.Number2
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 12345
							cell.Formatting = XlNumberFormat.NumberWithThousandSeparator
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 0.33
							cell.Formatting = XlNumberFormat.Percentage
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlNumberFormat.ShortDate
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlNumberFormat.ShortTime12
						End Using
					End Using
					' Custom formats
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Custom formats:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 4310.45
							cell.Formatting = New XlCellFormatting()
							cell.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 3426.75
							cell.Formatting = New XlCellFormatting()
							cell.Formatting.NumberFormat = "_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "" - ""??_-;_-@_-"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 0.333
							cell.Formatting = New XlCellFormatting()
							cell.Formatting.NumberFormat = "0.0%"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = New XlCellFormatting()
							cell.Formatting.NumberFormat = "dddd, mmmm d, yyyy"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 0.6234
							cell.Formatting = New XlCellFormatting()
							cell.Formatting.NumberFormat = "# ???/???"
						End Using
					End Using

					' .NET number formats
					sheet.SkipRows(1)
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = ".NET number formats"
							cell.Formatting = XlCellFormatting.Heading4
						End Using
					End Using
					' Standard format strings
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Standard formats:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 123.45
							cell.Formatting = XlCellFormatting.FromNetFormat("D", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 12345
							cell.Formatting = XlCellFormatting.FromNetFormat("E", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 0.33
							cell.Formatting = XlCellFormatting.FromNetFormat("P", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlCellFormatting.FromNetFormat("d", True)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlCellFormatting.FromNetFormat("t", True)
						End Using
					End Using
					' Custom format strings
					Using row As IXlRow = sheet.CreateRow()
						Using cell As IXlCell = row.CreateCell()
							cell.Value = "Custom formats:"
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 123.45
							cell.Formatting = XlCellFormatting.FromNetFormat("#0.00", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 12345
							cell.Formatting = XlCellFormatting.FromNetFormat("0.0##e+00", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = 0.333
							cell.Formatting = XlCellFormatting.FromNetFormat("Max={0:#.0%}", False)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlCellFormatting.FromNetFormat("dd-MM-yyyy", True)
						End Using
						Using cell As IXlCell = row.CreateCell()
							cell.Value = DateTime.Now
							cell.Formatting = XlCellFormatting.FromNetFormat("hh:mm tt", True)
						End Using
					End Using
				End Using
			End Using

'			#End Region ' #NumberFormat
		End Sub
	End Class
End Namespace