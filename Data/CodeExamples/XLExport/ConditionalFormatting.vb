Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports DevExpress.Export.Xl
Imports DevExpress.Spreadsheet

Namespace XLExportExamples
	Public NotInheritable Class ConditionalFormatting

		Private Sub New()
		End Sub
		Private Shared Sub Average(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#AverageRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = i + 1
								End Using
							Next j
						End Using
					Next i

					' Highlight values above average
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)) ' A1:A11
					Dim rule As New XlCondFmtRuleAboveAverage()
					rule.Condition = XlCondFmtAverageCondition.Above
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					' Highlight values above or equal to average
					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)) ' B1:B11
					rule = New XlCondFmtRuleAboveAverage()
					rule.Condition = XlCondFmtAverageCondition.AboveOrEqual
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					' Highlight values below average
					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)) ' C1:C11
					rule = New XlCondFmtRuleAboveAverage()
					rule.Condition = XlCondFmtAverageCondition.Below
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					' Highlight values below or equal to average
					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)) ' D1:D11
					rule = New XlCondFmtRuleAboveAverage()
					rule.Condition = XlCondFmtAverageCondition.BelowOrEqual
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #AverageRule
		End Sub

		Private Shared Sub CellIs(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#CellIsRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = i + 1
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = 12 - i
							End Using
						End Using
					Next i

					' Highlight values using "cell is" rules with "less than", "between" and "greater than" operators
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)) ' A1:A11
					Dim rule As New XlCondFmtRuleCellIs()
					rule.Operator = XlCondFmtOperator.LessThan
					rule.Formatting = XlCellFormatting.Bad
					rule.Value = 5
					formatting.Rules.Add(rule)
					rule = New XlCondFmtRuleCellIs()
					rule.Operator = XlCondFmtOperator.Between
					rule.Formatting = XlCellFormatting.Neutral
					rule.Value = 5
					rule.SecondValue = 8
					formatting.Rules.Add(rule)
					rule = New XlCondFmtRuleCellIs()
					rule.Operator = XlCondFmtOperator.GreaterThan
					rule.Formatting = XlCellFormatting.Good
					rule.Value = 8
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					' Highlight values using "cell is" rule with criteria value specified by formula
					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)) ' B1:B11
					rule = New XlCondFmtRuleCellIs()
					rule.Operator = XlCondFmtOperator.GreaterThan
					rule.Formatting = XlCellFormatting.Bad
					rule.Value = "=$A1+3"
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #CellIsRule
		End Sub

		Private Shared Sub Blanks(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#BlanksRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 9
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								If (i Mod 2) = 0 Then
									cell.Value = i + 1
								End If
							End Using
						End Using
					Next i

					' Format cells with blanks/no blanks
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9)) ' A1:A10
					' Highlight blank cells
					Dim rule As New XlCondFmtRuleBlanks(True)
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					' Highlight non blank cells
					rule = New XlCondFmtRuleBlanks(False) ' non blank cells
					rule.Formatting = XlCellFormatting.Good
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #BlanksRule
		End Sub

		Private Shared Sub Duplicates(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#DuplicatesRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = cell.ColumnIndex * cell.RowIndex + cell.RowIndex + 1
								End Using
							Next j
						End Using
					Next i

					' Highlight duplicate/unique values
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 10)) ' A1:D11
					' Highlight duplicate values
					formatting.Rules.Add(New XlCondFmtRuleDuplicates() With {.Formatting = XlCellFormatting.Bad})
					' Highlight unique values
					formatting.Rules.Add(New XlCondFmtRuleUnique() With {.Formatting = XlCellFormatting.Good})
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #DuplicatesRule
		End Sub

		Private Shared Sub Expression(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#ExpressionRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat, New XlFormulaParser())

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					Dim width() As Integer = { 80, 150, 90 }
					For i As Integer = 0 To 2
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = width(i)
							If i = 2 Then
								column.Formatting = New XlCellFormatting()
								column.Formatting.NumberFormat = "[$$-409] #,##0.00"
							End If
						End Using
					Next i
					Dim columnNames() As String = { "Account ID", "User Name", "Balance" }
					Using row As IXlRow = sheet.CreateRow()
						Dim headerRowFormatting As New XlCellFormatting()
						headerRowFormatting.Font = XlFont.BodyFont()
						headerRowFormatting.Font.Bold = True
						headerRowFormatting.Border = New XlBorder()
						headerRowFormatting.Border.BottomColor = Color.Black
						headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin
						For i As Integer = 0 To 2
							Using cell As IXlCell = row.CreateCell()
								cell.Value = columnNames(i)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using
					Dim accountIds() As String = { "A105", "A114", "B013", "C231", "D101", "D105" }
					Dim users() As String = { "Berry Dafoe", "Chris Cadwell", "Esta Mangold", "Liam Bell", "Simon Newman", "Wendy Underwood" }
					Dim balance() As Integer = { 155, 250, 48, 350, -15, 10 }
					For i As Integer = 0 To 5
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = accountIds(i)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = users(i)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = balance(i)
							End Using
						End Using
					Next i

					' Highlight values using expression rule
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 1, 2, 6)) ' A2:C7
					Dim rule As New XlCondFmtRuleExpression("AND($C2>0,$C2<50)")
					rule.Formatting = XlFill.SolidFill(Color.FromArgb(&Hff, &Hff, &Hcc))
					formatting.Rules.Add(rule)
					rule = New XlCondFmtRuleExpression("$C2<=0")
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #ExpressionRule
		End Sub

		Private Shared Sub SpecificText(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#SpecificTextRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					Dim width() As Integer = { 250, 180, 100 }
					For i As Integer = 0 To 2
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = width(i)
							If i = 2 Then
								column.Formatting = New XlCellFormatting()
								column.Formatting.NumberFormat = "_([$$-409]* #,##0.00_);_([$$-409]* \(#,##0.00\);_([$$-409]* ""-""??_);_(@_)"
							End If
						End Using
					Next i
					Dim columnNames() As String = { "Product", "Delivery", "Sales" }
					Using row As IXlRow = sheet.CreateRow()
						Dim headerRowFormatting As New XlCellFormatting()
						headerRowFormatting.Font = XlFont.BodyFont()
						headerRowFormatting.Font.Bold = True
						headerRowFormatting.Border = New XlBorder()
						headerRowFormatting.Border.BottomColor = Color.Black
						headerRowFormatting.Border.BottomLineStyle = XlBorderLineStyle.Thin
						For i As Integer = 0 To 2
							Using cell As IXlCell = row.CreateCell()
								cell.Value = columnNames(i)
								cell.ApplyFormatting(headerRowFormatting)
							End Using
						Next i
					End Using
					Dim products() As String = { "Camembert Pierrot", "Gorgonzola Telino", "Mascarpone Fabioli", "Mozzarella di Giovanni", "Queso Cabrales", "Raclette Courdavault" }
					Dim deliveries() As String = { "USA", "Worldwide", "USA", "Ships worldwide", "Worldwide except EU", "EU" }
					Dim sales() As Integer = { 15500, 20250, 12634, 35010, 15234, 10050 }
					For i As Integer = 0 To 5
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = products(i)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = deliveries(i)
							End Using
							Using cell As IXlCell = row.CreateCell()
								cell.Value = sales(i)
							End Using
						End Using
					Next i

					' Highlight values using specific text rule
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 1, 1, 6)) ' B2:B7
					Dim rule As New XlCondFmtRuleSpecificText(XlCondFmtSpecificTextType.Contains, "worldwide")
					rule.Formatting = XlCellFormatting.Neutral
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #SpecificTextRule
		End Sub

		Private Shared Sub TimePeriod(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#TimePeriodRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					Using column As IXlColumn = sheet.CreateColumn()
						column.WidthInPixels = 100
						column.ApplyFormatting(XlNumberFormat.ShortDate)
					End Using
					For i As Integer = 0 To 9
						Using row As IXlRow = sheet.CreateRow()
							Using cell As IXlCell = row.CreateCell()
								cell.Value = DateTime.Now.AddDays(row.RowIndex - 5)
							End Using
						End Using
					Next i

					' Highlight values using time period rules
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 9)) ' A1:A10
					Dim rule As New XlCondFmtRuleTimePeriod()
					rule.TimePeriod = XlCondFmtTimePeriod.Yesterday
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					rule = New XlCondFmtRuleTimePeriod()
					rule.TimePeriod = XlCondFmtTimePeriod.Today
					rule.Formatting = XlCellFormatting.Good
					formatting.Rules.Add(rule)
					rule = New XlCondFmtRuleTimePeriod()
					rule.TimePeriod = XlCondFmtTimePeriod.Tomorrow
					rule.Formatting = XlCellFormatting.Neutral
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #TimePeriodRule
		End Sub

		Private Shared Sub Top10(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#Top/BottomRules"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 9
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = cell.ColumnIndex * 4 + cell.RowIndex + 1
								End Using
							Next j
						End Using
					Next i

					' Highlight values using top/bottom rules
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 3, 9)) ' A1:D10
					' Bottom 10
					Dim rule As New XlCondFmtRuleTop10()
					rule.Bottom = True
					rule.Formatting = XlCellFormatting.Bad
					formatting.Rules.Add(rule)
					' Top 10
					rule = New XlCondFmtRuleTop10()
					rule.Formatting = XlCellFormatting.Good
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #Top/BottomRules
		End Sub

		Private Shared Sub DataBar(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#DataBarRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 2
						Using column As IXlColumn = sheet.CreateColumn()
							column.WidthInPixels = 100
						End Using
					Next i
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 2
								Using cell As IXlCell = row.CreateCell()
									Dim rowIndex As Integer = cell.RowIndex
									Dim columnIndex As Integer = cell.ColumnIndex
									If columnIndex = 0 Then
										cell.Value = rowIndex + 1
									ElseIf columnIndex = 1 Then
										cell.Value = rowIndex - 5
									Else
										cell.Value = If((rowIndex < 5), rowIndex + 1, 11 - rowIndex)
									End If
								End Using
							Next j
						End Using
					Next i

					' Format values using data bar rule
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)) ' A1:A11
					Dim rule As New XlCondFmtRuleDataBar()
					rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.2)
					rule.GradientFill = False
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)) ' B1:B11
					rule = New XlCondFmtRuleDataBar()
					rule.FillColor = Color.Green
					rule.BorderColor = Color.Green
					rule.AxisColor = Color.Brown
					rule.GradientFill = True
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)) ' C1:C11
					rule = New XlCondFmtRuleDataBar()
					rule.FillColor = XlColor.FromTheme(XlThemeColor.Accent4, 0.2)
					rule.MinLength = 10
					rule.MaxLength = 90
					rule.MinValue.ObjectType = XlCondFmtValueObjectType.Number
					rule.MinValue.Value = 3
					rule.Direction = XlDataBarDirection.RightToLeft
					rule.ShowValues = False
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #DataBarRule
		End Sub

		Private Shared Sub IconSet(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#IconSetRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									If cell.ColumnIndex Mod 2 = 0 Then
										cell.Value = cell.RowIndex + 1
									Else
										cell.Value = cell.RowIndex - 5
									End If
								End Using
							Next j
						End Using
					Next i

					' Format values using icon set rule
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)) ' A1:A11
					Dim rule As New XlCondFmtRuleIconSet()
					rule.IconSetType = XlCondFmtIconSetType.Arrows3
					rule.Priority = 1
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)) ' B1:B11
					rule = New XlCondFmtRuleIconSet()
					rule.IconSetType = XlCondFmtIconSetType.Flags3
					rule.Priority = 2
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)) ' C1:C11
					rule = New XlCondFmtRuleIconSet()
					rule.IconSetType = XlCondFmtIconSetType.Rating5
					rule.ShowValues = False
					rule.Priority = 3
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)) ' D1:D11
					rule = New XlCondFmtRuleIconSet()
					rule.IconSetType = XlCondFmtIconSetType.TrafficLights4
					rule.Reverse = True
					rule.Priority = 4
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #IconSetRule
		End Sub

		Private Shared Sub ColorScale(ByVal stream As Stream, ByVal documentFormat As XlDocumentFormat)
'			#Region "#ColorScaleRule"
			' Create exporter
			Dim exporter As IXlExporter = XlExport.CreateExporter(documentFormat)

			' Create document
			Using document As IXlDocument = exporter.CreateDocument(stream)
				document.Options.Culture = CultureInfo.CurrentCulture

				' Create worksheet
				Using sheet As IXlSheet = document.CreateSheet()

					' Create sample content
					For i As Integer = 0 To 10
						Using row As IXlRow = sheet.CreateRow()
							For j As Integer = 0 To 3
								Using cell As IXlCell = row.CreateCell()
									cell.Value = cell.RowIndex + 1
								End Using
							Next j
						End Using
					Next i

					' Format values using color scale rule
					Dim formatting As New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(0, 0, 0, 10)) ' A1:A11
					formatting.Ranges.Add(XlCellRange.FromLTRB(2, 0, 2, 10)) ' C1:C11
					Dim rule As New XlCondFmtRuleColorScale()
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)

					formatting = New XlConditionalFormatting()
					formatting.Ranges.Add(XlCellRange.FromLTRB(1, 0, 1, 10)) ' B1:B11
					formatting.Ranges.Add(XlCellRange.FromLTRB(3, 0, 3, 10)) ' D1:D11
					rule = New XlCondFmtRuleColorScale()
					rule.ColorScaleType = XlCondFmtColorScaleType.ColorScale2
					rule.MinColor = XlColor.FromTheme(XlThemeColor.Light1, 0.0)
					rule.MaxColor = XlColor.FromTheme(XlThemeColor.Accent1, 0.5)
					formatting.Rules.Add(rule)
					sheet.ConditionalFormattings.Add(formatting)
				End Using
			End Using

'			#End Region ' #ColorScaleRule
		End Sub

	End Class
End Namespace