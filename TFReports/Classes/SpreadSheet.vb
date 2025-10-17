Public Class SpreadSheet
    Event ProgressCounter(CounterMin As Integer, CounterMax As Integer, CounterValue As Integer)
    Private Const E28_ReasonACTIONED = "ACTIONED"
    Private Const E28_ReasonPOSTPONE = "POSTPONE"
    Private Const E28_ReasonREJECT = "REJECT"
    Private Structure TotalRowProps
        Dim TotalRowIndex As Integer
        Dim FirstDataRow As Integer
    End Structure
    Private Structure WorkSheetProps
        Dim WorkSheetName As String
        Dim RowCountWS As Integer
        Dim TotalsWS(,) As Double
        Dim PrevTotalKey1 As String
        Dim PrevTotalKey2 As String
        Dim GrandTotalRowsWS() As TotalRowProps
        Dim AirlineTotalRowsWS() As TotalRowProps
        Dim TypeTotalRowsWS() As TotalRowProps
        Dim FirstRowWS As Integer
        Dim LastRowWS As Integer
    End Structure
    Private Structure WorkbookProps
        Dim WorkBook As SpreadsheetLight.SLDocument
        Dim WorkSheets() As WorkSheetProps
        Dim WorkbookName As String
        Dim RowCountWB As Integer
        Dim TotalsWB(,) As Double
        Dim FirstRowWB As Integer
        Dim LastRowWB As Integer
    End Structure

    Private ThisDocument As New SpreadsheetLight.SLDocument
    Private RowCount As Integer
    Private ReadOnly xlDecStyle As SpreadsheetLight.SLStyle
    Private ReadOnly xlStyleLightGray As SpreadsheetLight.SLStyle
    Private ReadOnly mdsDataSet As DataSet
    Private ReadOnly mReportTitle As String
    Private ReadOnly MonthNames As String() = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}

    Public Sub New(pDs As DataSet, ByVal ReportTitle As String)

        mReportTitle = ReportTitle
        mdsDataSet = New DataSet
        mdsDataSet = pDs

        ThisDocument = New SpreadsheetLight.SLDocument
        xlDecStyle = ThisDocument.CreateStyle
        xlDecStyle.FormatCode = "#,##0.00;-#,##0.00;"

        xlStyleLightGray = ThisDocument.CreateStyle
        xlStyleLightGray.Font.Bold = True
        xlStyleLightGray.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleLightGray.Fill.SetPatternForegroundColor(Color.LightGray)
    End Sub
    'Ye Olde Spreadesheetes
    'Public Sub E00_Euronav(ByVal FileName As String)

    '    Dim xlWorkSheet As New SpreadsheetLight.SLDocument

    '    Dim xlINVCount As Integer = 0
    '    Dim xlCNSCount As Integer = 0

    '    Try
    '        xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "INV")
    '        With xlWorkSheet
    '            .FreezePanes(1, 0)
    '            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
    '            xlNumStyle.FormatCode = "@"
    '            .SetColumnStyle(1, 14, xlNumStyle)
    '            xlNumStyle.FormatCode = "dd/mm/yyyy"
    '            .SetColumnStyle(3, xlNumStyle)
    '            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
    '            .SetColumnStyle(7, xlNumStyle)
    '            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
    '            Next
    '        End With
    '        xlWorkSheet.AddWorksheet("CNS")
    '        With xlWorkSheet
    '            .FreezePanes(1, 0)
    '            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
    '            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
    '            .SetColumnStyle(5, 11, xlNumStyle)
    '            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
    '            Next
    '        End With

    '        For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
    '            If mdsDataSet.Tables(0).Rows(i).Item(0).ToString.StartsWith("C") Then
    '                xlWorkSheet.SelectWorksheet("CNS")
    '                xlCNSCount += 1
    '                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                    xlWorkSheet.SetCellValue(xlCNSCount + 1, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
    '                Next
    '            Else
    '                xlWorkSheet.SelectWorksheet("INV")
    '                xlINVCount += 1
    '                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                    xlWorkSheet.SetCellValue(xlINVCount + 1, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
    '                Next
    '            End If
    '        Next
    '        xlWorkSheet.SaveAs(FileName)
    '        MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    'Public Sub E02_BSPMonthReportbyairline(ByVal mBSPMonthDate As String, ByVal FileName As String)

    '    Dim xlWorkSheet As New SpreadsheetLight.SLDocument
    '    Dim mPrevSheetName As String = ""

    '    Dim xlINVCount As Integer = 0
    '    Dim xlCNSCount As Integer = 0

    '    Try
    '        xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, mBSPMonthDate)
    '        With xlWorkSheet
    '            .FreezePanes(1, 0)
    '            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
    '            xlNumStyle.FormatCode = "@"
    '            .SetColumnStyle(1, 5, xlNumStyle)
    '            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
    '            .SetColumnStyle(6, 12, xlNumStyle)
    '            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
    '            Next
    '        End With
    '        xlWorkSheet.AddWorksheet("Totals")
    '        With xlWorkSheet
    '            .FreezePanes(1, 0)
    '            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
    '            xlNumStyle.FormatCode = "@"
    '            .SetColumnStyle(1, 3, xlNumStyle)
    '            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
    '            .SetColumnStyle(4, 10, xlNumStyle)
    '            .SetCellValue(1, 1, mdsDataSet.Tables(0).Columns(0).Caption)
    '            .SetCellValue(1, 2, mdsDataSet.Tables(0).Columns(1).Caption)
    '            .SetCellValue(1, 3, mdsDataSet.Tables(0).Columns(2).Caption)
    '            For j = 5 To mdsDataSet.Tables(0).Columns.Count - 1
    '                .SetCellValue(1, j - 1, mdsDataSet.Tables(0).Columns(j).Caption)
    '            Next
    '        End With
    '        mPrevSheetName = "Totals"
    '        ' BSP Month Report by airline
    '        Dim xlNumStyleTot As SpreadsheetLight.SLStyle = xlWorkSheet.CreateStyle
    '        xlNumStyleTot.SetFontBold(True)
    '        With xlWorkSheet
    '            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
    '                RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
    '                If mdsDataSet.Tables(0).Rows(i).Item(3).ToString = "" Then
    '                    If mPrevSheetName <> "Totals" Then
    '                        .SelectWorksheet("Totals")
    '                        mPrevSheetName = "Totals"
    '                    End If
    '                    xlCNSCount += 1
    '                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                        If j <= 2 Then
    '                            .SetCellValue(xlCNSCount + 1, j + 1, IIf(IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(j)), "", mdsDataSet.Tables(0).Rows(i).Item(j)))
    '                        ElseIf j >= 5 Then
    '                            .SetCellValue(xlCNSCount + 1, j - 1, IIf(IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(j)), "", mdsDataSet.Tables(0).Rows(i).Item(j)))
    '                        End If
    '                    Next
    '                End If
    '                If mPrevSheetName <> mBSPMonthDate Then
    '                    .SelectWorksheet(mBSPMonthDate)
    '                    mPrevSheetName = mBSPMonthDate
    '                End If
    '                xlINVCount += 1
    '                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                    .SetCellValue(xlINVCount + 1, j + 1, IIf(IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(j)), "", mdsDataSet.Tables(0).Rows(i).Item(j)))
    '                Next
    '                If mdsDataSet.Tables(0).Rows(i).Item(3) = "" Then
    '                    .SetRowStyle(xlINVCount + 1, xlNumStyleTot)
    '                End If
    '            Next
    '            .SelectWorksheet("Totals")
    '            .SetRowStyle(xlCNSCount + 1, xlNumStyleTot)
    '            .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)
    '            .SelectWorksheet(mBSPMonthDate)
    '            .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)
    '            .SaveAs(FileName)
    '        End With

    '        MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub
    'Public Sub E03_BSPFortnightReportbyticket_Original(ByVal mBSPFortDate As Date, ByVal FileName As String)

    '    Dim pobjWorkBook As New SpreadsheetLight.SLDocument
    '    Dim pISheetName As String = Format(mBSPFortDate, "yyyyMMdd" & " I")
    '    Dim pDSheetName As String = Format(mBSPFortDate, "yyyyMMdd" & " D")
    '    Dim pPrevSheetName As String = ""
    '    Dim pIRowCount As Integer = 0
    '    Dim pDRowCount As Integer = 0

    '    ' Prepare styles for formatting totals rows
    '    Dim xlStyleGrandTotal As SpreadsheetLight.SLStyle = pobjWorkBook.CreateStyle
    '    Dim xlStyleAirlineTotal As SpreadsheetLight.SLStyle = pobjWorkBook.CreateStyle
    '    Dim xlStyleTypeTotal As SpreadsheetLight.SLStyle = pobjWorkBook.CreateStyle

    '    ' 0 for grand total
    '    xlStyleGrandTotal.Font.Bold = True
    '    xlStyleGrandTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
    '    xlStyleGrandTotal.Fill.SetPatternForegroundColor(Color.Silver)

    '    ' 1 for airline total
    '    xlStyleAirlineTotal.Font.Bold = True
    '    xlStyleAirlineTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
    '    xlStyleAirlineTotal.Fill.SetPatternForegroundColor(Color.Aqua)

    '    ' 2 for type total
    '    xlStyleTypeTotal.Font.Bold = True
    '    xlStyleTypeTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
    '    xlStyleTypeTotal.Fill.SetPatternForegroundColor(Color.Orange)

    '    Try

    '        ' BSP Fortnight Report by ticket

    '        Dim pITotals(2, 10) As Decimal ' 0=grand total 1=type total 2=air total
    '        Dim pITotalsGrand(0) As TotalRowProps
    '        Dim pITotalsType(0) As TotalRowProps
    '        Dim pITotalsAirline(0) As TotalRowProps
    '        Dim pIPrevAirline As String = ""
    '        Dim pIPrevType As String = ""

    '        Dim pDTotals(2, 10) As Decimal
    '        Dim pDTotalsGrand(0) As TotalRowProps
    '        Dim pDTotalsType(0) As TotalRowProps
    '        Dim pDTotalsAirline(0) As TotalRowProps
    '        Dim pDPrevAirline As String = ""
    '        Dim pDPrevType As String = ""

    '        ' initialize arrays which keep track of which rows in excel have totals and must be formatted accordingly at the end

    '        For ii As Integer = pITotals.GetLowerBound(0) To pITotals.GetUpperBound(0)
    '            For jj As Integer = pITotals.GetUpperBound(1) To pITotals.GetUpperBound(1)
    '                pITotals(ii, jj) = 0
    '                pDTotals(ii, jj) = 0
    '            Next
    '        Next
    '        ' redim the row indices so as to be ready for the next totals group coming up
    '        ReDim pDTotalsAirline(1)
    '        pDTotalsAirline(1).FirstDataRow = -1
    '        ReDim pDTotalsType(1)
    '        pDTotalsType(1).FirstDataRow = -1
    '        ReDim pDTotalsGrand(1)
    '        pDTotalsGrand(1).FirstDataRow = -1

    '        ReDim pITotalsAirline(1)
    '        pITotalsAirline(1).FirstDataRow = -1
    '        ReDim pITotalsType(1)
    '        pITotalsAirline(1).FirstDataRow = -1
    '        ReDim pITotalsGrand(1)
    '        pITotalsAirline(1).FirstDataRow = -1

    '        ' prepare excel

    '        pobjWorkBook.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, pISheetName)
    '        E03_PrepareWorksheet(pobjWorkBook)
    '        pobjWorkBook.AddWorksheet(pDSheetName)
    '        E03_PrepareWorksheet(pobjWorkBook)
    '        pPrevSheetName = pDSheetName


    '        ' loop through data grid, row by row
    '        For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
    '            RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
    '            ' domestic tickets (D in column 2) go into a separate worksheet
    '            If mdsDataSet.Tables(0).Rows(i).Item(1).ToString = "D" Then
    '                If pPrevSheetName <> pDSheetName Then
    '                    pobjWorkBook.SelectWorksheet(pDSheetName)
    '                    pPrevSheetName = pDSheetName
    '                End If
    '                ' check for subtotals triggered by change of airline/ticket type
    '                If pDPrevAirline <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Or pDPrevType <> mdsDataSet.Tables(0).Rows(i).Item(2).ToString Then
    '                    ' if previous ticket type is not empty then this is not the first row
    '                    If pDPrevType <> "" Then
    '                        ' add a total row for type total
    '                        pDRowCount += 1
    '                        pobjWorkBook.SetCellValue(pDRowCount + 1, 1, pDPrevAirline)
    '                        pobjWorkBook.SetCellValue(pDRowCount + 1, 3, pDPrevType)
    '                        pobjWorkBook.SetCellValue(pDRowCount + 1, 4, "TOTAL")
    '                        ' enter values for type total and accumulate values for airline total
    '                        ' for columns with values (7-16)
    '                        For jj As Integer = 7 To 16
    '                            pobjWorkBook.SetCellValue(pDRowCount + 1, jj + 1, pDTotals(2, jj - 7))
    '                            pDTotals(1, jj - 7) += pDTotals(2, jj - 7)
    '                            pDTotals(2, jj - 7) = 0
    '                        Next
    '                        ' store row number which has type total
    '                        pDTotalsAirline(pDTotalsAirline.GetUpperBound(0)).TotalRowIndex = pDRowCount
    '                        ' and then redim the index array for the next totals group
    '                        ReDim Preserve pDTotalsAirline(pDTotalsAirline.GetUpperBound(0) + 1)
    '                        pDTotalsAirline(pDTotalsAirline.GetUpperBound(0)).FirstDataRow = -1
    '                        ' check for subtotal triggered by change of airline
    '                        If pDPrevAirline <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Then
    '                            ' if previous airline is not empty then this is not the first row
    '                            If pDPrevAirline <> "" Then
    '                                ' add a total row in the data array for airline total
    '                                pDRowCount += 1
    '                                pobjWorkBook.SetCellValue(pDRowCount + 1, 1, pDPrevAirline)
    '                                pobjWorkBook.SetCellValue(pDRowCount + 1, 4, "TOTAL")
    '                                ' enter values for airline total and accumulate values for grand total
    '                                ' for columns with values (7-16)
    '                                For jj As Integer = 7 To 16
    '                                    pobjWorkBook.SetCellValue(pDRowCount + 1, jj + 1, pDTotals(1, jj - 7))
    '                                    pDTotals(0, jj - 7) += pDTotals(1, jj - 7)
    '                                    pDTotals(1, jj - 7) = 0
    '                                Next
    '                                ' store row number which has airline total
    '                                pDTotalsType(pDTotalsType.GetUpperBound(0)).TotalRowIndex = pDRowCount
    '                                ReDim Preserve pDTotalsType(pDTotalsType.GetUpperBound(0) + 1)
    '                                pDTotalsType(pDTotalsType.GetUpperBound(0)).FirstDataRow = -1
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '                ' add a row in the data array for the data
    '                pDRowCount += 1
    '                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                    pobjWorkBook.SetCellValue(pDRowCount + 1, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
    '                    ' accumulate ticket values to type total except for columns which are percentages (11 and 13)
    '                    If j >= 7 And j <= 16 And j <> 11 And j <> 13 Then
    '                        pDTotals(2, j - 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(j))
    '                    End If
    '                    If pDTotalsAirline(pDTotalsAirline.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pDTotalsAirline(pDTotalsAirline.GetUpperBound(0)).FirstDataRow = pDRowCount
    '                    End If
    '                    If pDTotalsType(pDTotalsType.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pDTotalsType(pDTotalsType.GetUpperBound(0)).FirstDataRow = pDRowCount
    '                    End If
    '                    If pDTotalsGrand(pDTotalsGrand.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pDTotalsGrand(pDTotalsGrand.GetUpperBound(0)).FirstDataRow = pDRowCount
    '                    End If
    '                Next
    '                ' store this row's airline and type for comparison to trigger subtotals
    '                pDPrevAirline = mdsDataSet.Tables(0).Rows(i).Item(0).ToString
    '                pDPrevType = mdsDataSet.Tables(0).Rows(i).Item(2).ToString
    '            Else
    '                ' non domestic tickets are processed here (I in column 2 - actually not D) go into a spearate worksheet
    '                ' the process is identical to the above half of the if statement and should be encapsulated 
    '                ' TODO Encapsulate the two different worksheets
    '                If pPrevSheetName <> pISheetName Then
    '                    pobjWorkBook.SelectWorksheet(pISheetName)
    '                    pPrevSheetName = pISheetName
    '                End If
    '                If pIPrevAirline <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Or pIPrevType <> mdsDataSet.Tables(0).Rows(i).Item(2).ToString Then
    '                    If pIPrevType <> "" Then
    '                        pIRowCount += 1
    '                        pobjWorkBook.SetCellValue(pIRowCount + 1, 1, pIPrevAirline)
    '                        pobjWorkBook.SetCellValue(pIRowCount + 1, 3, pIPrevType)
    '                        pobjWorkBook.SetCellValue(pIRowCount + 1, 4, "TOTAL")
    '                        For jj As Integer = 7 To 16
    '                            pobjWorkBook.SetCellValue(pIRowCount + 1, jj + 1, pITotals(2, jj - 7))
    '                            pITotals(1, jj - 7) += pITotals(2, jj - 7)
    '                            pITotals(2, jj - 7) = 0
    '                        Next
    '                        pITotalsAirline(pITotalsAirline.GetUpperBound(0)).TotalRowIndex = pIRowCount
    '                        ReDim Preserve pITotalsAirline(pITotalsAirline.GetUpperBound(0) + 1)
    '                        pITotalsAirline(pITotalsAirline.GetUpperBound(0)).FirstDataRow = -1
    '                    End If
    '                    If pIPrevAirline <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Then
    '                        If pIPrevAirline <> "" Then
    '                            pIRowCount += 1
    '                            pobjWorkBook.SetCellValue(pIRowCount + 1, 1, pIPrevAirline)
    '                            pobjWorkBook.SetCellValue(pIRowCount + 1, 4, "TOTAL")
    '                            For jj As Integer = 7 To 16
    '                                pobjWorkBook.SetCellValue(pIRowCount + 1, jj + 1, pITotals(1, jj - 7))
    '                                pITotals(0, jj - 7) += pITotals(1, jj - 7)
    '                                pITotals(1, jj - 7) = 0
    '                            Next
    '                            pITotalsType(pITotalsType.GetUpperBound(0)).TotalRowIndex = pIRowCount
    '                            ReDim Preserve pITotalsType(pITotalsType.GetUpperBound(0) + 1)
    '                            pITotalsType(pITotalsType.GetUpperBound(0)).FirstDataRow = -1
    '                        End If
    '                    End If
    '                End If
    '                pIRowCount += 1
    '                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                    pobjWorkBook.SetCellValue(pIRowCount + 1, j + 1, IIf(IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(j)), "", mdsDataSet.Tables(0).Rows(i).Item(j)))
    '                    If j >= 7 And j <= 16 And j <> 11 And j <> 13 Then
    '                        pITotals(2, j - 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(j))
    '                    End If
    '                    If pITotalsAirline(pITotalsAirline.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pITotalsAirline(pITotalsAirline.GetUpperBound(0)).FirstDataRow = pIRowCount
    '                    End If
    '                    If pITotalsType(pITotalsType.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pITotalsType(pITotalsType.GetUpperBound(0)).FirstDataRow = pIRowCount
    '                    End If
    '                    If pITotalsGrand(pITotalsGrand.GetUpperBound(0)).FirstDataRow = -1 Then
    '                        pITotalsGrand(pITotalsGrand.GetUpperBound(0)).FirstDataRow = pIRowCount
    '                    End If

    '                Next
    '                pIPrevAirline = mdsDataSet.Tables(0).Rows(i).Item(0).ToString
    '                pIPrevType = mdsDataSet.Tables(0).Rows(i).Item(2).ToString
    '            End If
    '        Next

    '        ' END OF MAIN LOOP WHICH LOOPS THROUGH DATAGRID ROWS

    '        ' Add type and airline totals for the last tickets in worksheet D
    '        With pobjWorkBook
    '            .SelectWorksheet(pDSheetName)
    '            pDRowCount += 1
    '            .SetCellValue(pDRowCount + 1, 1, pDPrevAirline)
    '            .SetCellValue(pDRowCount + 1, 3, pDPrevType)
    '            .SetCellValue(pDRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pDRowCount + 1, jj + 1, pDTotals(2, jj - 7))
    '                pDTotals(1, jj - 7) += pDTotals(2, jj - 7)
    '                pDTotals(2, jj - 7) = 0
    '            Next
    '            pDTotalsAirline(pDTotalsAirline.GetUpperBound(0)).TotalRowIndex = pDRowCount
    '            pDRowCount += 1
    '            .SetCellValue(pDRowCount + 1, 1, pDPrevAirline)
    '            .SetCellValue(pDRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pDRowCount + 1, jj + 1, pDTotals(1, jj - 7))
    '                pDTotals(0, jj - 7) += pDTotals(1, jj - 7)
    '                pDTotals(1, jj - 7) = 0
    '            Next
    '            pDTotalsType(pDTotalsType.GetUpperBound(0)).TotalRowIndex = pDRowCount
    '            pDRowCount += 1
    '            .SetCellValue(pDRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pDRowCount + 1, jj + 1, pDTotals(0, jj - 7))
    '            Next
    '            pDTotalsGrand(pDTotalsGrand.GetUpperBound(0)).TotalRowIndex = pDRowCount

    '            ' loop through arrays of row numbers which have totals and format accordingly
    '            For ii = 1 To pDTotalsGrand.GetUpperBound(0)
    '                .SetCellStyle(pDTotalsGrand(ii).TotalRowIndex + 1, 1, pDTotalsGrand(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleGrandTotal)
    '                .GroupRows(pDTotalsGrand(ii).FirstDataRow + 1, pDTotalsGrand(ii).TotalRowIndex)
    '            Next ii
    '            For ii = 1 To pDTotalsType.GetUpperBound(0)
    '                .SetCellStyle(pDTotalsType(ii).TotalRowIndex + 1, 1, pDTotalsType(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleAirlineTotal)
    '                .GroupRows(pDTotalsType(ii).FirstDataRow + 1, pDTotalsType(ii).TotalRowIndex)
    '            Next
    '            For ii = 1 To pDTotalsAirline.GetUpperBound(0)
    '                .SetCellStyle(pDTotalsAirline(ii).TotalRowIndex + 1, 1, pDTotalsAirline(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleTypeTotal)
    '                .GroupRows(pDTotalsAirline(ii).FirstDataRow + 1, pDTotalsAirline(ii).TotalRowIndex)
    '            Next

    '            .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)

    '            ' Add type and airline totals for the last tickets in worksheet I
    '            .SelectWorksheet(pISheetName)
    '            pIRowCount += 1
    '            .SetCellValue(pIRowCount + 1, 1, pIPrevAirline)
    '            .SetCellValue(pIRowCount + 1, 3, pIPrevType)
    '            .SetCellValue(pIRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pIRowCount + 1, jj + 1, pITotals(2, jj - 7))
    '                pITotals(1, jj - 7) += pITotals(2, jj - 7)
    '                pITotals(2, jj - 7) = 0
    '            Next
    '            pITotalsAirline(pITotalsAirline.GetUpperBound(0)).TotalRowIndex = pIRowCount
    '            pIRowCount += 1
    '            .SetCellValue(pIRowCount + 1, 1, pIPrevAirline)
    '            .SetCellValue(pIRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pIRowCount + 1, jj + 1, pITotals(1, jj - 7))
    '                pITotals(0, jj - 7) += pITotals(1, jj - 7)
    '                pITotals(1, jj - 7) = 0
    '            Next
    '            pITotalsType(pITotalsType.GetUpperBound(0)).TotalRowIndex = pIRowCount
    '            pIRowCount += 1
    '            .SetCellValue(pIRowCount + 1, 4, "TOTAL")
    '            For jj As Integer = 7 To 16
    '                .SetCellValue(pIRowCount + 1, jj + 1, pITotals(0, jj - 7))
    '            Next
    '            pITotalsGrand(pITotalsGrand.GetUpperBound(0)).TotalRowIndex = pIRowCount

    '            ' loop through arrays of row numbers which have totals and format accordingly
    '            For ii = 1 To pITotalsGrand.GetUpperBound(0)
    '                .SetCellStyle(pITotalsGrand(ii).TotalRowIndex + 1, 1, pITotalsGrand(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count - 1, xlStyleGrandTotal)
    '                .GroupRows(pITotalsGrand(ii).FirstDataRow + 1, pITotalsGrand(ii).TotalRowIndex)
    '            Next ii
    '            For ii = 1 To pITotalsType.GetUpperBound(0)
    '                .SetCellStyle(pITotalsType(ii).TotalRowIndex + 1, 1, pITotalsType(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count - 1, xlStyleAirlineTotal)
    '                .GroupRows(pITotalsType(ii).FirstDataRow + 1, pITotalsType(ii).TotalRowIndex)
    '            Next
    '            For ii = 1 To pITotalsAirline.GetUpperBound(0)
    '                .SetCellStyle(pITotalsAirline(ii).TotalRowIndex + 1, 1, pITotalsAirline(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count - 1, xlStyleTypeTotal)
    '                .GroupRows(pITotalsAirline(ii).FirstDataRow + 1, pITotalsAirline(ii).TotalRowIndex)
    '            Next
    '            .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)

    '            .SaveAs(FileName)

    '        End With

    '        MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Public Sub E03_BSPFortnightReportbyticket(ByVal mBSPFortDate As Date, ByVal FileName As String)

        Dim pobjWorkBook As New WorkbookProps
        Dim pPrevSheetName As String = ""

        Try

            ' BSP Fortnight Report by ticket
            ' 0=grand total 1=type total 2=air total

            With pobjWorkBook
                ReDim .WorkSheets(1)
                .WorkSheets(0).WorkSheetName = Format(mBSPFortDate, "yyyyMMdd" & " I")
                .WorkSheets(1).WorkSheetName = Format(mBSPFortDate, "yyyyMMdd" & " D")
                For iSheet As Integer = 0 To 1
                    With .WorkSheets(iSheet)
                        .RowCountWS = 0
                        ReDim .TotalsWS(2, 10)
                        For ii As Integer = .TotalsWS.GetLowerBound(iSheet) To .TotalsWS.GetUpperBound(0)
                            For jj As Integer = .TotalsWS.GetUpperBound(1) To .TotalsWS.GetUpperBound(1)
                                .TotalsWS(ii, jj) = 0
                                .TotalsWS(ii, jj) = 0
                            Next
                        Next
                        .PrevTotalKey1 = ""
                        .PrevTotalKey2 = ""
                        ' redim the row indices so as to be ready for the next totals group coming up
                        ReDim .AirlineTotalRowsWS(1)
                        .AirlineTotalRowsWS(1).FirstDataRow = -1
                        ReDim .TypeTotalRowsWS(1)
                        .TypeTotalRowsWS(1).FirstDataRow = -1
                        ReDim .GrandTotalRowsWS(1)
                        .GrandTotalRowsWS(1).FirstDataRow = -1
                    End With
                Next

                .WorkBook = New SpreadsheetLight.SLDocument
                .WorkBook.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, .WorkSheets(0).WorkSheetName)
                E03_PrepareWorksheet(.WorkBook)
                .WorkBook.AddWorksheet(.WorkSheets(1).WorkSheetName)
                E03_PrepareWorksheet(.WorkBook)
                pPrevSheetName = .WorkSheets(1).WorkSheetName
            End With



            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                ' domestic tickets (D in column 2) go into a separate worksheet
                If mdsDataSet.Tables(0).Rows(i).Item(1).ToString = "D" Then
                    If pPrevSheetName <> pobjWorkBook.WorkSheets(1).WorkSheetName Then
                        pobjWorkBook.WorkBook.SelectWorksheet(pobjWorkBook.WorkSheets(1).WorkSheetName)
                        pPrevSheetName = pobjWorkBook.WorkSheets(1).WorkSheetName
                    End If
                    E03_AddDataRows(pobjWorkBook, 1, mdsDataSet.Tables(0).Rows(i))
                Else
                    ' non domestic tickets are processed here (I in column 2 - actually not D) go into a spearate worksheet
                    If pPrevSheetName <> pobjWorkBook.WorkSheets(0).WorkSheetName Then
                        pobjWorkBook.WorkBook.SelectWorksheet(pobjWorkBook.WorkSheets(0).WorkSheetName)
                        pPrevSheetName = pobjWorkBook.WorkSheets(0).WorkSheetName
                    End If
                    E03_AddDataRows(pobjWorkBook, 0, mdsDataSet.Tables(0).Rows(i))
                End If
            Next

            ' Add type and airline totals for the last tickets in worksheet D
            E03_AddClosingTotalRows(pobjWorkBook, 1)
            ' Add type and airline totals for the last tickets in worksheet I
            E03_AddClosingTotalRows(pobjWorkBook, 0)

            pobjWorkBook.WorkBook.SaveAs(FileName)

            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub E03_AddClosingTotalRows(ByRef pobjWorkBook As WorkbookProps, ByVal WorksheetIndex As Integer)

        ' Prepare styles for formatting totals rows
        Dim xlStyleGrandTotal As SpreadsheetLight.SLStyle = pobjWorkBook.WorkBook.CreateStyle
        Dim xlStyleAirlineTotal As SpreadsheetLight.SLStyle = pobjWorkBook.WorkBook.CreateStyle
        Dim xlStyleTypeTotal As SpreadsheetLight.SLStyle = pobjWorkBook.WorkBook.CreateStyle

        ' for grand total
        xlStyleGrandTotal.Font.Bold = True
        xlStyleGrandTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleGrandTotal.Fill.SetPatternForegroundColor(Color.Silver)

        ' for airline total
        xlStyleAirlineTotal.Font.Bold = True
        xlStyleAirlineTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleAirlineTotal.Fill.SetPatternForegroundColor(Color.Aqua)

        ' for type total
        xlStyleTypeTotal.Font.Bold = True
        xlStyleTypeTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleTypeTotal.Fill.SetPatternForegroundColor(Color.Orange)

        With pobjWorkBook.WorkSheets(WorksheetIndex)
            ' Add type and airline totals for the last tickets in worksheet D
            pobjWorkBook.WorkBook.SelectWorksheet(.WorkSheetName)
            .RowCountWS += 1
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 1, .PrevTotalKey1)
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 3, .PrevTotalKey2)
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 4, "TOTAL")
            For jj As Integer = 7 To 16
                pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, jj + 1, .TotalsWS(2, jj - 7))
                .TotalsWS(1, jj - 7) += .TotalsWS(2, jj - 7)
                .TotalsWS(2, jj - 7) = 0
            Next
            .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0)).TotalRowIndex = .RowCountWS
            .RowCountWS += 1
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 1, .PrevTotalKey1)
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 4, "TOTAL")
            For jj As Integer = 7 To 16
                pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, jj + 1, .TotalsWS(1, jj - 7))
                .TotalsWS(0, jj - 7) += .TotalsWS(1, jj - 7)
                .TotalsWS(1, jj - 7) = 0
            Next
            .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0)).TotalRowIndex = .RowCountWS
            .RowCountWS += 1
            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 4, "TOTAL")
            For jj As Integer = 7 To 16
                pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, jj + 1, .TotalsWS(0, jj - 7))
            Next
            .GrandTotalRowsWS(.GrandTotalRowsWS.GetUpperBound(0)).TotalRowIndex = .RowCountWS

            ' loop through arrays of row numbers which have totals and format accordingly
            For ii = 1 To .GrandTotalRowsWS.GetUpperBound(0)
                pobjWorkBook.WorkBook.SetCellStyle(.GrandTotalRowsWS(ii).TotalRowIndex + 1, 1, .GrandTotalRowsWS(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleGrandTotal)
                pobjWorkBook.WorkBook.GroupRows(.GrandTotalRowsWS(ii).FirstDataRow + 1, .GrandTotalRowsWS(ii).TotalRowIndex)
            Next ii
            For ii = 1 To .AirlineTotalRowsWS.GetUpperBound(0)
                pobjWorkBook.WorkBook.SetCellStyle(.AirlineTotalRowsWS(ii).TotalRowIndex + 1, 1, .AirlineTotalRowsWS(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleAirlineTotal)
                pobjWorkBook.WorkBook.GroupRows(.AirlineTotalRowsWS(ii).FirstDataRow + 1, .AirlineTotalRowsWS(ii).TotalRowIndex)
            Next
            For ii = 1 To .TypeTotalRowsWS.GetUpperBound(0)
                pobjWorkBook.WorkBook.SetCellStyle(.TypeTotalRowsWS(ii).TotalRowIndex + 1, 1, .TypeTotalRowsWS(ii).TotalRowIndex + 1, mdsDataSet.Tables(0).Columns.Count, xlStyleTypeTotal)
                pobjWorkBook.WorkBook.GroupRows(.TypeTotalRowsWS(ii).FirstDataRow + 1, .TypeTotalRowsWS(ii).TotalRowIndex)
            Next

            pobjWorkBook.WorkBook.AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)
        End With

    End Sub
    Private Sub E03_AddDataRows(ByRef pobjWorkBook As WorkbookProps, ByVal ActiveWorksheet As Integer, ByVal DataRow As DataRow)

        With pobjWorkBook.WorkSheets(ActiveWorksheet)

            ' check for subtotals triggered by change of airline/ticket type
            If .PrevTotalKey1 <> DataRow.Item(0).ToString Or .PrevTotalKey2 <> DataRow.Item(2).ToString Then
                ' if previous ticket type is not empty then this is not the first row
                If .PrevTotalKey2 <> "" Then
                    ' add a total row for type total
                    .RowCountWS += 1
                    pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 1, .PrevTotalKey1)
                    pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 3, .PrevTotalKey2)
                    pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 4, "TOTAL")
                    ' enter values for type total and accumulate values for airline total
                    ' for columns with values (7-16)
                    For jj As Integer = 7 To 16
                        pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, jj + 1, .TotalsWS(2, jj - 7))
                        .TotalsWS(1, jj - 7) += .TotalsWS(2, jj - 7)
                        .TotalsWS(2, jj - 7) = 0
                    Next
                    ' store row number which has type total
                    .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0)).TotalRowIndex = .RowCountWS
                    ' and then redim the index array for the next totals group
                    ReDim Preserve .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0) + 1)
                    .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0)).FirstDataRow = -1
                    ' check for subtotal triggered by change of airline
                    If .PrevTotalKey1 <> DataRow.Item(0).ToString Then
                        ' if previous airline is not empty then this is not the first row
                        If .PrevTotalKey1 <> "" Then
                            ' add a total row in the data array for airline total
                            .RowCountWS += 1
                            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 1, .PrevTotalKey1)
                            pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, 4, "TOTAL")
                            ' enter values for airline total and accumulate values for grand total
                            ' for columns with values (7-16)
                            For jj As Integer = 7 To 16
                                pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, jj + 1, .TotalsWS(1, jj - 7))
                                .TotalsWS(0, jj - 7) += .TotalsWS(1, jj - 7)
                                .TotalsWS(1, jj - 7) = 0
                            Next
                            ' store row number which has airline total
                            .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0)).TotalRowIndex = .RowCountWS
                            ReDim Preserve .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0) + 1)
                            .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0)).FirstDataRow = -1
                        End If
                    End If
                End If
            End If
            ' add a row in the data array for the data
            .RowCountWS += 1
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                pobjWorkBook.WorkBook.SetCellValue(.RowCountWS + 1, j + 1, DataRow.Item(j))
                ' accumulate ticket values to type total except for columns which are percentages (11 and 13)
                If j >= 7 And j <= 16 And j <> 11 And j <> 13 Then
                    .TotalsWS(2, j - 7) += CDec(DataRow.Item(j))
                End If
                If .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0)).FirstDataRow = -1 Then
                    .AirlineTotalRowsWS(.AirlineTotalRowsWS.GetUpperBound(0)).FirstDataRow = .RowCountWS
                End If
                If .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0)).FirstDataRow = -1 Then
                    .TypeTotalRowsWS(.TypeTotalRowsWS.GetUpperBound(0)).FirstDataRow = .RowCountWS
                End If
                If .GrandTotalRowsWS(.GrandTotalRowsWS.GetUpperBound(0)).FirstDataRow = -1 Then
                    .GrandTotalRowsWS(.GrandTotalRowsWS.GetUpperBound(0)).FirstDataRow = .RowCountWS
                End If
            Next
            ' store this row's airline and type for comparison to trigger subtotals
            .PrevTotalKey1 = DataRow.Item(0).ToString
            .PrevTotalKey2 = DataRow.Item(2).ToString
        End With

    End Sub
    Private Sub E03_PrepareWorksheet(ByRef xlWorkSheet As SpreadsheetLight.SLDocument)

        With xlWorkSheet
            .FreezePanes(1, 0)
            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
            xlNumStyle.FormatCode = "@"
            .SetColumnStyle(1, 4, xlNumStyle)
            .SetColumnStyle(6, 7, xlNumStyle)
            .SetColumnStyle(18, xlNumStyle)
            xlNumStyle.FormatCode = "dd/mm/yyyy"
            .SetColumnStyle(5, xlNumStyle)
            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
            .SetColumnStyle(8, 17, xlNumStyle)
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
            Next
        End With

    End Sub
    Public Sub E04_TicketInfo(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim pobjWorkBook As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try
            Dim pTickList As String = ""
            For i As Integer = 0 To mReport.TextEntryItemsCount
                If mReport.TextEntryItems(i).Length = 10 Then
                    pTickList &= "|" & mReport.TextEntryItems(i) & "|"
                End If
            Next
            ' TicketInfo
            With pobjWorkBook

                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "TicketInfo")
                .FreezePanes(1, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 13, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(4, xlNumStyle)
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        .SetCellValue(xlINVCount + 1, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    pTickList = pTickList.Replace("|" & mdsDataSet.Tables(0).Rows(i).Item(0) & "|", "")
                Next
                If pTickList <> "" Then
                    xlINVCount += 1
                    .SetCellValue(xlINVCount + 1, 1, "Tickets not in Travel Force")
                    .SetCellValue(xlINVCount + 1, 2, pTickList)
                End If
                .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)
                .SaveAs(FileName)
            End With
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Public Sub E05_ClientTurnover(ByVal FileName As String)

        Dim pobjWorkBook As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            ' Client Turnover
            With pobjWorkBook

                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Client Turnover")
                .FreezePanes(1, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 3, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(4, xlNumStyle)
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(1, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        .SetCellValue(xlINVCount + 1, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                Next
                .AutoFitColumn(0, mdsDataSet.Tables(0).Columns.Count)
                .SaveAs(FileName)
            End With
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E06_ProfitPerClientwithBudgetComparison(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String) ' ByVal mDateFrom As Date, ByVal mReport.DateTo As Date, ByVal mReport.FromYTD As Date, ByVal mReport.ToYTD As Date, ByVal mReport.FromPY As Date, ByVal mReport.ToPY As Date, ByVal FileName As String)
        '  0-GroupName	
        '  1-Client	
        '  2-Sales	
        '  3-Profit	
        '  4-Pax	
        '  5-ProfitPerPax ------	
        '  6-SalesYTD	
        '  7-ProfitYTD	
        '  8-PaxYTD	
        '  9-ProfitPerPaxYTD ------	
        ' 10-SalesPYTD	
        ' 11-ProfitPYTD	
        ' 12-PaxPYTD	
        ' 13-ProfitPerPaxPYTD -----	
        ' 14-SalesPYCurr	
        ' 15-ProfitPYCurr	
        ' 16-PaxPYCurr	
        ' 17-ProfitPerPaxPYCurr	----
        ' 18-SalesBudgetCurr	
        ' 19-ProfitBudgetCurr	
        ' 20-PaxBudgetCurr	
        ' 21-ProfitPerPaxBudgetCurr ----	
        ' 22-SalesBudgetYTD	
        ' 23-ProfitBudgetYTD	
        ' 24-PaxBudgetYTD	
        ' 25-ProfitPerPaxBudgetYTD ----

        Dim pobjWorkbooks(0) As WorkbookProps

        Try

            ' Profit Per client group with PY
            With pobjWorkbooks(0)
                ReDim .WorkSheets(3)
                .WorkSheets(0).WorkSheetName = "Sales"
                .WorkSheets(1).WorkSheetName = "Profit"
                .WorkSheets(2).WorkSheetName = "Pax"
                .WorkSheets(3).WorkSheetName = "ProfitPerPax"
                For iSheet As Integer = 0 To 3
                    With .WorkSheets(iSheet)
                        .RowCountWS = 0
                        ReDim .TotalsWS(1, 25)
                        For ii As Integer = .TotalsWS.GetLowerBound(0) To .TotalsWS.GetUpperBound(0)
                            For jj As Integer = .TotalsWS.GetUpperBound(1) To .TotalsWS.GetUpperBound(1)
                                .TotalsWS(ii, jj) = 0
                                .TotalsWS(ii, jj) = 0
                            Next
                        Next
                    End With
                Next
                .WorkBook = New SpreadsheetLight.SLDocument
                .WorkBook.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, .WorkSheets(0).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(1).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(2).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(3).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
            End With

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)

                If pobjWorkbooks.GetUpperBound(0) = 0 Then
                    pobjWorkbooks(0).WorkBook.SelectWorksheet(pobjWorkbooks(0).WorkSheets(0).WorkSheetName)
                End If

                ' enter new data row in master workbook
                With pobjWorkbooks(0)
                    .WorkSheets(0).RowCountWS += 1

                    '  0-GroupName	
                    '  1-Client	

                    '  2-Sales	
                    ' 18-SalesBudgetCurr
                    '    Comparison 2/18
                    '  6-SalesYTD	
                    ' 22-SalesBudgetYTD	
                    '    Comparison 6/22
                    ' 14-SalesPYCurr	
                    '    Comparison 2/14
                    ' 10-SalesPYTD	
                    '    Comparison 6/10

                    '  3-Profit	
                    ' 19-ProfitBudgetCurr	
                    '    Comparison
                    '  7-ProfitYTD	
                    ' 23-ProfitBudgetYTD	
                    '    Comparison
                    ' 15-ProfitPYCurr	
                    '    Comparison
                    ' 11-ProfitPYTD	
                    '    Comparison

                    '  4-Pax	
                    ' 20-PaxBudgetCurr	
                    '    Comparison
                    '  8-PaxYTD	
                    ' 24-PaxBudgetYTD	
                    '    Comparison
                    ' 16-PaxPYCurr	
                    '    Comparison
                    ' 12-PaxPYTD	
                    '    Comparison

                    '  5-ProfitPerPax ------	
                    ' 21-ProfitPerPaxBudgetCurr ----	
                    '    Comparison
                    '  9-ProfitPerPaxYTD ------	
                    ' 25-ProfitPerPaxBudgetYTD ----
                    '    Comparison
                    ' 17-ProfitPerPaxPYCurr	----
                    '    Comparison
                    ' 13-ProfitPerPaxPYTD -----	
                    '    Comparison
                    For iWS As Integer = 0 To 2
                        .WorkBook.SelectWorksheet(.WorkSheets(iWS).WorkSheetName)
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 18), mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 22), mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 9, 10, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 14), mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 11, 12, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 10), mReport.BooleanOption1)

                        .WorkSheets(iWS).TotalsWS(0, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                        .WorkSheets(iWS).TotalsWS(1, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                        .WorkSheets(iWS).TotalsWS(0, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                        .WorkSheets(iWS).TotalsWS(1, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                        .WorkSheets(iWS).TotalsWS(0, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                        .WorkSheets(iWS).TotalsWS(1, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                        .WorkSheets(iWS).TotalsWS(0, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                        .WorkSheets(iWS).TotalsWS(1, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                        .WorkSheets(iWS).TotalsWS(0, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                        .WorkSheets(iWS).TotalsWS(1, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                        .WorkSheets(iWS).TotalsWS(0, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                        .WorkSheets(iWS).TotalsWS(1, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                    Next
                    ' Profit per Pax
                    Dim pPPPCurr As Decimal
                    Dim pPPPCurrBud As Decimal
                    Dim pPPPytd As Decimal
                    Dim pPPPytdBud As Decimal
                    Dim pPPPPrev As Decimal
                    Dim pPPPPrevytd As Decimal
                    pPPPCurr = DivNums(mdsDataSet.Tables(0).Rows(i).Item(3), mdsDataSet.Tables(0).Rows(i).Item(4))
                    pPPPCurrBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(19), mdsDataSet.Tables(0).Rows(i).Item(20))
                    pPPPytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(7), mdsDataSet.Tables(0).Rows(i).Item(8))
                    pPPPytdBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(23), mdsDataSet.Tables(0).Rows(i).Item(24))
                    pPPPPrev = DivNums(mdsDataSet.Tables(0).Rows(i).Item(15), mdsDataSet.Tables(0).Rows(i).Item(16))
                    pPPPPrevytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(11), mdsDataSet.Tables(0).Rows(i).Item(12))
                    .WorkBook.SelectWorksheet(.WorkSheets(3).WorkSheetName)
                    .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                    .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                    E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, mReport.BooleanOption1)
                    E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, mReport.BooleanOption1)
                    E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, mReport.BooleanOption1)
                    E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, mReport.BooleanOption1)

                End With
            Next
            ' Grand Total
            E12_AddTotalsRows(pobjWorkbooks(0), 0, "Grand Total", mReport.BooleanOption1)

            Dim pAllFiles As String = ""
            With pobjWorkbooks(0).WorkBook
                For iWS = 0 To 3
                    pobjWorkbooks(0).WorkBook.SelectWorksheet(pobjWorkbooks(0).WorkSheets(iWS).WorkSheetName)
                    .AutoFitColumn(0, 12)
                Next
                E12_SetAreaColours(pobjWorkbooks(0))
                .SaveAs(FileName)
            End With

            pAllFiles &= FileName & vbCrLf
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Public Sub E07_ProfitPerOPSgroup(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            ' Profit Per OPS group

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group")
                .FreezePanes(3, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(4, 8, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 3, xlNumStyle)
                .SetCellStyle(1, 4, xlNumStyle)
                .SetCellStyle(1, 5, xlNumStyle)
                xlWorkSheet.SetCellValue(1, 4, Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy"))
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(3, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next
            End With

            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    xlWorkSheet.SetCellValue(xlINVCount + 3, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                Next
            Next

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E08_ProfitPerclientgroup(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            ' Profit Per client group
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group")
                .FreezePanes(3, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 7, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(6, xlNumStyle)
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 2, xlNumStyle)
                .SetCellStyle(1, 4, xlNumStyle)
                .SetCellStyle(1, 5, xlNumStyle)
                xlWorkSheet.SetCellValue(1, 4, Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy"))
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(3, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next
            End With

            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    xlWorkSheet.SetCellValue(xlINVCount + 3, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                Next
            Next

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E09_ProfitPerOpsGroupExtra(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal mPrevDateFrom As Date, ByVal mPrevDateTo As Date, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            ' Profit Per ops group extra

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group")
                .FreezePanes(3, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 12, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(13, xlNumStyle)
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(2, xlNumStyle)
                .SetCellStyle(1, 4, xlNumStyle)
                .SetCellStyle(1, 5, xlNumStyle)
                xlWorkSheet.SetCellValue(1, 4, Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy"))
                xlWorkSheet.SetCellValue(1, 5, Format(mPrevDateFrom, "dd/MM/yyyy") & "-" & Format(mPrevDateTo, "dd/MM/yyyy"))
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(3, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next
            End With

            xlINVCount = 0

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    xlWorkSheet.SetCellValue(xlINVCount + 3, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j).ToString)
                Next
            Next

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E10_ProfitPerClientGroupExtra(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            ' Profit Per client group extra

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Per Group")
                .FreezePanes(3, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 12, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(13, xlNumStyle)
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(2, xlNumStyle)
                .SetCellStyle(1, 4, xlNumStyle)
                .SetCellStyle(1, 5, xlNumStyle)
                xlWorkSheet.SetCellValue(1, 4, Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy"))
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(3, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next
            End With

            xlINVCount = 0

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    xlWorkSheet.SetCellValue(xlINVCount + 3, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j).ToString)
                Next
            Next
            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E12_ProfitPerclientgroupwithBudgetComparison(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String) ' ByVal mDateFrom As Date, ByVal mReport.DateTo As Date, ByVal mReport.FromYTD As Date, ByVal mReport.ToYTD As Date, ByVal mReport.FromPY As Date, ByVal mReport.ToPY As Date, ByVal FileName As String)
        '  0-GroupName	
        '  1-Client	
        '  2-Sales	
        '  3-Profit	
        '  4-Pax	
        '  5-ProfitPerPax ------	
        '  6-SalesYTD	
        '  7-ProfitYTD	
        '  8-PaxYTD	
        '  9-ProfitPerPaxYTD ------	
        ' 10-SalesPYTD	
        ' 11-ProfitPYTD	
        ' 12-PaxPYTD	
        ' 13-ProfitPerPaxPYTD -----	
        ' 14-SalesPYCurr	
        ' 15-ProfitPYCurr	
        ' 16-PaxPYCurr	
        ' 17-ProfitPerPaxPYCurr	----
        ' 18-SalesBudgetCurr	
        ' 19-ProfitBudgetCurr	
        ' 20-PaxBudgetCurr	
        ' 21-ProfitPerPaxBudgetCurr ----	
        ' 22-SalesBudgetYTD	
        ' 23-ProfitBudgetYTD	
        ' 24-PaxBudgetYTD	
        ' 25-ProfitPerPaxBudgetYTD ----

        Dim pobjWorkbooks(0) As WorkbookProps
        Dim pstrPrevGroupName As String = ""
        Try

            ' Profit Per client group with PY
            With pobjWorkbooks(0)
                ReDim .WorkSheets(3)
                .WorkSheets(0).WorkSheetName = "Sales"
                .WorkSheets(1).WorkSheetName = "Profit"
                .WorkSheets(2).WorkSheetName = "Pax"
                .WorkSheets(3).WorkSheetName = "ProfitPerPax"
                For iSheet As Integer = 0 To 3
                    With .WorkSheets(iSheet)
                        .RowCountWS = 0
                        ReDim .TotalsWS(1, 25)
                        For ii As Integer = .TotalsWS.GetLowerBound(0) To .TotalsWS.GetUpperBound(0)
                            For jj As Integer = .TotalsWS.GetUpperBound(1) To .TotalsWS.GetUpperBound(1)
                                .TotalsWS(ii, jj) = 0
                                .TotalsWS(ii, jj) = 0
                            Next
                        Next
                    End With
                Next
                .WorkBook = New SpreadsheetLight.SLDocument
                .WorkBook.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, .WorkSheets(0).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(1).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(2).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
                .WorkBook.AddWorksheet(.WorkSheets(3).WorkSheetName)
                E12_PrepareWorksheet(.WorkBook, mReport)
            End With

            Dim a1() As String = mReport.TextEntry.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
            Dim pGroups As String = ""
            For i = a1.GetLowerBound(0) To a1.GetUpperBound(0)
                pGroups &= Right("00" & a1(i), 2) & "|"
            Next
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                Dim OpsGroup As String = mdsDataSet.Tables(0).Rows(i).Item(0).ToString

                If pGroups = "" Or pGroups.IndexOf(OpsGroup.Substring(0, 2)) >= 0 Then

                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)

                    If pobjWorkbooks.GetUpperBound(0) = 0 OrElse pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)).WorkbookName <> OpsGroup Then
                        pobjWorkbooks(0).WorkBook.SelectWorksheet(pobjWorkbooks(0).WorkSheets(0).WorkSheetName)
                        ' If this is not the first group, then enter the totals of the previous group into the Master workbook
                        If pobjWorkbooks.GetUpperBound(0) > 0 Then
                            E12_AddTotalsRows(pobjWorkbooks(0), 1, pstrPrevGroupName & " TOTAL", mReport.BooleanOption1)
                        End If
                        ' save the last row of data for the previous group
                        pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)).WorkSheets(0).LastRowWS = pobjWorkbooks(0).WorkSheets(0).RowCountWS - 1

                        ' create a new workbook for the new group
                        ReDim Preserve pobjWorkbooks(pobjWorkbooks.GetUpperBound(0) + 1)
                        With pobjWorkbooks(pobjWorkbooks.GetUpperBound(0))
                            ReDim .WorkSheets(3)
                            .WorkSheets(0).WorkSheetName = "Sales"
                            .WorkSheets(1).WorkSheetName = "Profit"
                            .WorkSheets(2).WorkSheetName = "Pax"
                            .WorkSheets(3).WorkSheetName = "ProfitPerPax"
                            For iSheet As Integer = 0 To 3
                                With .WorkSheets(iSheet)
                                    .RowCountWS = 0
                                    ReDim .TotalsWS(1, 25)
                                    .FirstRowWS = pobjWorkbooks(0).WorkSheets(0).RowCountWS + 1
                                    .LastRowWS = -1
                                    For ii As Integer = .TotalsWS.GetLowerBound(0) To .TotalsWS.GetUpperBound(0)
                                        For jj As Integer = .TotalsWS.GetUpperBound(1) To .TotalsWS.GetUpperBound(1)
                                            .TotalsWS(ii, jj) = 0
                                            .TotalsWS(ii, jj) = 0
                                        Next
                                    Next
                                End With
                            Next
                            .WorkBook = New SpreadsheetLight.SLDocument
                            .WorkbookName = OpsGroup
                            .WorkBook.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, .WorkSheets(0).WorkSheetName)
                            E12_PrepareWorksheet(.WorkBook, mReport)
                            .WorkBook.AddWorksheet(.WorkSheets(1).WorkSheetName)
                            E12_PrepareWorksheet(.WorkBook, mReport)
                            .WorkBook.AddWorksheet(.WorkSheets(2).WorkSheetName)
                            E12_PrepareWorksheet(.WorkBook, mReport)
                            .WorkBook.AddWorksheet(.WorkSheets(3).WorkSheetName)
                            E12_PrepareWorksheet(.WorkBook, mReport)
                        End With
                    End If

                    ' enter new data row in master workbook
                    With pobjWorkbooks(0)
                        pstrPrevGroupName = OpsGroup
                        .WorkSheets(0).RowCountWS += 1

                        '  0-GroupName	
                        '  1-Client	

                        '  2-Sales	
                        ' 18-SalesBudgetCurr
                        '    Comparison 2/18
                        '  6-SalesYTD	
                        ' 22-SalesBudgetYTD	
                        '    Comparison 6/22
                        ' 14-SalesPYCurr	
                        '    Comparison 2/14
                        ' 10-SalesPYTD	
                        '    Comparison 6/10

                        '  3-Profit	
                        ' 19-ProfitBudgetCurr	
                        '    Comparison
                        '  7-ProfitYTD	
                        ' 23-ProfitBudgetYTD	
                        '    Comparison
                        ' 15-ProfitPYCurr	
                        '    Comparison
                        ' 11-ProfitPYTD	
                        '    Comparison

                        '  4-Pax	
                        ' 20-PaxBudgetCurr	
                        '    Comparison
                        '  8-PaxYTD	
                        ' 24-PaxBudgetYTD	
                        '    Comparison
                        ' 16-PaxPYCurr	
                        '    Comparison
                        ' 12-PaxPYTD	
                        '    Comparison

                        '  5-ProfitPerPax ------	
                        ' 21-ProfitPerPaxBudgetCurr ----	
                        '    Comparison
                        '  9-ProfitPerPaxYTD ------	
                        ' 25-ProfitPerPaxBudgetYTD ----
                        '    Comparison
                        ' 17-ProfitPerPaxPYCurr	----
                        '    Comparison
                        ' 13-ProfitPerPaxPYTD -----	
                        '    Comparison
                        For iWS As Integer = 0 To 2
                            .WorkBook.SelectWorksheet(.WorkSheets(iWS).WorkSheetName)
                            .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, OpsGroup)
                            .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                            E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 18), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 22), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 9, 10, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 14), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 11, 12, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 10), mReport.BooleanOption1)

                            .WorkSheets(iWS).TotalsWS(0, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                            .WorkSheets(iWS).TotalsWS(1, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                            .WorkSheets(iWS).TotalsWS(0, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                            .WorkSheets(iWS).TotalsWS(1, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                            .WorkSheets(iWS).TotalsWS(0, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                            .WorkSheets(iWS).TotalsWS(1, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                            .WorkSheets(iWS).TotalsWS(0, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                            .WorkSheets(iWS).TotalsWS(1, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                            .WorkSheets(iWS).TotalsWS(0, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                            .WorkSheets(iWS).TotalsWS(1, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                            .WorkSheets(iWS).TotalsWS(0, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                            .WorkSheets(iWS).TotalsWS(1, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                        Next

                        ' Profit per Pax
                        Dim pPPPCurr As Decimal
                        Dim pPPPCurrBud As Decimal
                        Dim pPPPytd As Decimal
                        Dim pPPPytdBud As Decimal
                        Dim pPPPPrev As Decimal
                        Dim pPPPPrevytd As Decimal

                        pPPPCurr = DivNums(mdsDataSet.Tables(0).Rows(i).Item(3), mdsDataSet.Tables(0).Rows(i).Item(4))
                        pPPPCurrBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(19), mdsDataSet.Tables(0).Rows(i).Item(20))
                        pPPPytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(7), mdsDataSet.Tables(0).Rows(i).Item(8))
                        pPPPytdBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(23), mdsDataSet.Tables(0).Rows(i).Item(24))
                        pPPPPrev = DivNums(mdsDataSet.Tables(0).Rows(i).Item(15), mdsDataSet.Tables(0).Rows(i).Item(16))
                        pPPPPrevytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(11), mdsDataSet.Tables(0).Rows(i).Item(12))
                        .WorkBook.SelectWorksheet(.WorkSheets(3).WorkSheetName)
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, OpsGroup)
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(0), .WorkSheets(0).RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, mReport.BooleanOption1)

                    End With

                    ' enter new data row in group's workbook
                    With pobjWorkbooks(pobjWorkbooks.GetUpperBound(0))
                        .WorkBook.SelectWorksheet(.WorkSheets(0).WorkSheetName)
                        .WorkSheets(0).RowCountWS += 1
                        For iWS As Integer = 0 To 2
                            .WorkBook.SelectWorksheet(.WorkSheets(iWS).WorkSheetName)
                            .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, OpsGroup)
                            .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                            E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 18), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 22), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 9, 10, mdsDataSet.Tables(0).Rows(i).Item(iWS + 2), mdsDataSet.Tables(0).Rows(i).Item(iWS + 14), mReport.BooleanOption1)
                            E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 11, 12, mdsDataSet.Tables(0).Rows(i).Item(iWS + 6), mdsDataSet.Tables(0).Rows(i).Item(iWS + 10), mReport.BooleanOption1)

                            .WorkSheets(iWS).TotalsWS(0, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                            .WorkSheets(iWS).TotalsWS(1, 3) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 2))
                            .WorkSheets(iWS).TotalsWS(0, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                            .WorkSheets(iWS).TotalsWS(1, 4) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 18))
                            .WorkSheets(iWS).TotalsWS(0, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                            .WorkSheets(iWS).TotalsWS(1, 5) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 6))
                            .WorkSheets(iWS).TotalsWS(0, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                            .WorkSheets(iWS).TotalsWS(1, 6) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 22))
                            .WorkSheets(iWS).TotalsWS(0, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                            .WorkSheets(iWS).TotalsWS(1, 7) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 14))
                            .WorkSheets(iWS).TotalsWS(0, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                            .WorkSheets(iWS).TotalsWS(1, 8) += CDec(mdsDataSet.Tables(0).Rows(i).Item(iWS + 10))
                        Next
                        ' Profit per Pax
                        Dim pPPPCurr As Decimal
                        Dim pPPPCurrBud As Decimal
                        Dim pPPPytd As Decimal
                        Dim pPPPytdBud As Decimal
                        Dim pPPPPrev As Decimal
                        Dim pPPPPrevytd As Decimal
                        pPPPCurr = DivNums(mdsDataSet.Tables(0).Rows(i).Item(3), mdsDataSet.Tables(0).Rows(i).Item(4))
                        pPPPCurrBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(19), mdsDataSet.Tables(0).Rows(i).Item(20))
                        pPPPytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(7), mdsDataSet.Tables(0).Rows(i).Item(8))
                        pPPPytdBud = DivNums(mdsDataSet.Tables(0).Rows(i).Item(23), mdsDataSet.Tables(0).Rows(i).Item(24))
                        pPPPPrev = DivNums(mdsDataSet.Tables(0).Rows(i).Item(15), mdsDataSet.Tables(0).Rows(i).Item(16))
                        pPPPPrevytd = DivNums(mdsDataSet.Tables(0).Rows(i).Item(11), mdsDataSet.Tables(0).Rows(i).Item(12))
                        .WorkBook.SelectWorksheet(.WorkSheets(3).WorkSheetName)
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, OpsGroup)
                        .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                        E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, mReport.BooleanOption1)
                        E12_AddRowValues(pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)), .WorkSheets(0).RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, mReport.BooleanOption1)
                    End With
                End If

            Next
            ' Group Total for last group when all data rows finished
            E12_AddTotalsRows(pobjWorkbooks(0), 1, pstrPrevGroupName & "TOTAL", mReport.BooleanOption1)
            ' save the last row of data for the previous group
            pobjWorkbooks(pobjWorkbooks.GetUpperBound(0)).WorkSheets(0).LastRowWS = pobjWorkbooks(0).WorkSheets(0).RowCountWS - 1

            ' Grand Total
            E12_AddTotalsRows(pobjWorkbooks(0), 0, "Grand Total", mReport.BooleanOption1)

            Dim pAllFiles As String = ""
            With pobjWorkbooks(0).WorkBook
                For iWS = 0 To 3
                    pobjWorkbooks(0).WorkBook.SelectWorksheet(pobjWorkbooks(0).WorkSheets(iWS).WorkSheetName)
                    ' group rows for grand total
                    .GroupRows(pobjWorkbooks(0).WorkSheets(0).FirstRowWS + 3, pobjWorkbooks(0).WorkSheets(0).RowCountWS + 2)
                    ' group rows for group totals
                    For i = 1 To pobjWorkbooks.GetUpperBound(0)
                        .GroupRows(pobjWorkbooks(i).WorkSheets(0).FirstRowWS + 3, pobjWorkbooks(i).WorkSheets(0).LastRowWS + 3)
                        .CollapseRows(pobjWorkbooks(i).WorkSheets(0).LastRowWS + 4)
                    Next
                    .AutoFitColumn(0, 12)
                Next
                E12_SetAreaColours(pobjWorkbooks(0))
                .SaveAs(FileName)
            End With

            pAllFiles &= FileName & vbCrLf
            For iSheet As Integer = 1 To pobjWorkbooks.GetUpperBound(0)
                E12_AddTotalsRows(pobjWorkbooks(iSheet), 0, "Total", mReport.BooleanOption1)
                For iWS = 0 To 3
                    pobjWorkbooks(iSheet).WorkBook.SelectWorksheet(pobjWorkbooks(iSheet).WorkSheets(iWS).WorkSheetName)
                    pobjWorkbooks(iSheet).WorkBook.AutoFitColumn(0, 12)
                Next
                Dim pFilename As String = FileName.Replace(".xlsx", "-" & pobjWorkbooks(iSheet).WorkbookName & ".xlsx")
                E12_SetAreaColours(pobjWorkbooks(iSheet))
                pobjWorkbooks(iSheet).WorkBook.SaveAs(pFilename)
                pAllFiles &= pFilename & vbCrLf
            Next
            MessageBox.Show("File(s) saved: " & pAllFiles, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Function DivNums(ByVal Val1 As Object, ByVal Val2 As Object) As Decimal
        Try
            If Val2 = 0 Then
                Return 0
            Else
                Return Val1 / Val2
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Private Sub E12_AddRowValues(ByRef pWorkBook As WorkbookProps, ByVal Row As Integer, ByVal FirstCell As Integer, ByVal SecondCell As Integer, ByVal ThirdCell As Integer, ByVal Value1 As Object, ByVal Value2 As Object, ByVal DiffAsPercentage As Boolean)

        With pWorkBook
            .WorkBook.SetCellValue(Row, FirstCell, Value1)
            .WorkBook.SetCellValue(Row, SecondCell, Value2)
            If Value1 * Value2 <> 0 Then
                If DiffAsPercentage Then
                    .WorkBook.SetCellValue(Row, ThirdCell, Value1 / Value2 * 100)
                Else
                    .WorkBook.SetCellValue(Row, ThirdCell, Value1 - Value2)
                End If
                If Value1 < Value2 Then
                    Dim xlStyle As SpreadsheetLight.SLStyle = .WorkBook.CreateStyle
                    xlStyle.Font.FontColor = Color.Red
                    .WorkBook.SetCellStyle(Row, ThirdCell, xlStyle)
                End If
            End If
        End With

    End Sub
    Private Sub E12_AddRowValues(ByRef pWorkBook As WorkbookProps, ByVal Row As Integer, ByVal FirstCell As Integer, ByVal SecondCell As Integer, ByVal Value1 As Object, ByVal Value2 As Object, ByVal DiffAsPercentage As Boolean)

        With pWorkBook
            .WorkBook.SetCellValue(Row, FirstCell, Value2)
            If Value1 * Value2 <> 0 Then
                If DiffAsPercentage Then
                    .WorkBook.SetCellValue(Row, SecondCell, Value1 / Value2 * 100)
                Else
                    .WorkBook.SetCellValue(Row, SecondCell, Value1 - Value2)
                End If
                If Value1 < Value2 Then
                    Dim xlStyle As SpreadsheetLight.SLStyle = .WorkBook.CreateStyle
                    xlStyle.Font.FontColor = Color.Red
                    .WorkBook.SetCellStyle(Row, SecondCell, xlStyle)
                End If
            End If
        End With

    End Sub
    Private Sub E12_SetAreaColours(ByRef pWorkBook As WorkbookProps)

        For iWS = 0 To 3
            pWorkBook.WorkBook.SelectWorksheet(pWorkBook.WorkSheets(iWS).WorkSheetName)
            Dim xlStyle As SpreadsheetLight.SLStyle = pWorkBook.WorkBook.CreateStyle
            xlStyle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
            xlStyle.Fill.SetPatternForegroundColor(Color.GreenYellow)
            pWorkBook.WorkBook.SetCellStyle(1, 3, pWorkBook.WorkSheets(0).RowCountWS + 3, 5, xlStyle)
            xlStyle.Fill.SetPatternForegroundColor(Color.SpringGreen)
            pWorkBook.WorkBook.SetCellStyle(1, 6, pWorkBook.WorkSheets(0).RowCountWS + 3, 8, xlStyle)
            xlStyle.Fill.SetPatternForegroundColor(Color.PowderBlue)
            pWorkBook.WorkBook.SetCellStyle(1, 9, pWorkBook.WorkSheets(0).RowCountWS + 3, 10, xlStyle)
            xlStyle.Fill.SetPatternForegroundColor(Color.SkyBlue)
            pWorkBook.WorkBook.SetCellStyle(1, 11, pWorkBook.WorkSheets(0).RowCountWS + 3, 12, xlStyle)
        Next

    End Sub
    Private Sub E12_AddTotalsRows(ByRef pWorkBook As WorkbookProps, ByVal TotalsLevel As Integer, ByVal TotalsDescription As String, ByVal DiffAsPercentage As Boolean)

        With pWorkBook
            Dim xlNumStyleB As SpreadsheetLight.SLStyle = .WorkBook.CreateStyle
            xlNumStyleB.Font.Bold = True
            .WorkSheets(0).RowCountWS += 1
            For iWS = 0 To 2
                .WorkBook.SelectWorksheet(.WorkSheets(iWS).WorkSheetName)
                .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, TotalsDescription)
                E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 3, 4, 5, .WorkSheets(iWS).TotalsWS(TotalsLevel, 3), .WorkSheets(iWS).TotalsWS(TotalsLevel, 4), DiffAsPercentage)
                E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 6, 7, 8, .WorkSheets(iWS).TotalsWS(TotalsLevel, 5), .WorkSheets(iWS).TotalsWS(TotalsLevel, 6), DiffAsPercentage)
                E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 9, 10, .WorkSheets(iWS).TotalsWS(TotalsLevel, 3), .WorkSheets(iWS).TotalsWS(TotalsLevel, 7), DiffAsPercentage)
                E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 11, 12, .WorkSheets(iWS).TotalsWS(TotalsLevel, 5), .WorkSheets(iWS).TotalsWS(TotalsLevel, 8), DiffAsPercentage)
                .WorkBook.SetCellStyle(.WorkSheets(0).RowCountWS + 3, 1, .WorkSheets(0).RowCountWS + 3, 12, xlNumStyleB)
            Next
            ' Profit per Pax
            Dim pPPPCurr As Decimal
            Dim pPPPCurrBud As Decimal
            Dim pPPPytd As Decimal
            Dim pPPPytdBud As Decimal
            Dim pPPPPrev As Decimal
            Dim pPPPPrevytd As Decimal
            pPPPCurr = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 3), .WorkSheets(2).TotalsWS(TotalsLevel, 3))
            pPPPCurrBud = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 4), .WorkSheets(2).TotalsWS(TotalsLevel, 4))
            pPPPytd = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 5), .WorkSheets(2).TotalsWS(TotalsLevel, 5))
            pPPPytdBud = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 6), .WorkSheets(2).TotalsWS(TotalsLevel, 6))
            pPPPPrev = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 7), .WorkSheets(2).TotalsWS(TotalsLevel, 7))
            pPPPPrevytd = DivNums(.WorkSheets(1).TotalsWS(TotalsLevel, 8), .WorkSheets(2).TotalsWS(TotalsLevel, 8))
            .WorkBook.SelectWorksheet(.WorkSheets(3).WorkSheetName)
            .WorkBook.SetCellValue(.WorkSheets(0).RowCountWS + 3, 1, TotalsDescription)
            E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 3, 4, 5, pPPPCurr, pPPPCurrBud, DiffAsPercentage)
            E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 6, 7, 8, pPPPytd, pPPPytdBud, DiffAsPercentage)
            E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 9, 10, pPPPCurr, pPPPPrev, DiffAsPercentage)
            E12_AddRowValues(pWorkBook, .WorkSheets(0).RowCountWS + 3, 11, 12, pPPPytd, pPPPPrevytd, DiffAsPercentage)
            .WorkBook.SetCellStyle(.WorkSheets(0).RowCountWS + 3, 1, .WorkSheets(0).RowCountWS + 3, 12, xlNumStyleB)
            For iWS = 0 To 2
                For j = 0 To 8
                    .WorkSheets(iWS).TotalsWS(TotalsLevel, j) = 0
                Next
            Next
        End With

    End Sub
    Private Sub E12_PrepareWorksheet(ByRef xlWorkSheet As SpreadsheetLight.SLDocument, ByRef mReport As Reports.ReportsCollection)

        With xlWorkSheet
            .FreezePanes(3, 0)
            Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

            xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
            .SetColumnStyle(3, 12, xlNumStyle)
            xlNumStyle.Font.Bold = True
            xlNumStyle.FormatCode = "@"
            .SetColumnStyle(1, 2, xlNumStyle)
            xlNumStyle.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center
            .SetCellStyle(1, 3, xlNumStyle)
            .SetCellValue(1, 3, Format(mReport.E12_FromCurr, "dd/MM/yyyy") & "-" & Format(mReport.E12_ToCurr, "dd/MM/yyyy"))
            .MergeWorksheetCells(1, 3, 1, 5)
            .SetCellStyle(2, 3, xlNumStyle)
            .SetCellValue(2, 3, "Current Month")
            .MergeWorksheetCells(2, 3, 2, 5)
            .SetCellStyle(1, 6, xlNumStyle)
            .SetCellValue(1, 6, Format(mReport.E12_FromYTD, "dd/MM/yyyy") & "-" & Format(mReport.E12_ToYTD, "dd/MM/yyyy"))
            .MergeWorksheetCells(1, 6, 1, 8)
            .SetCellStyle(2, 6, xlNumStyle)
            .SetCellValue(2, 6, "Current Year To Date")
            .MergeWorksheetCells(2, 6, 2, 8)
            .SetCellStyle(1, 9, xlNumStyle)
            .SetCellValue(1, 9, Format(mReport.E12_FromPYCurr, "dd/MM/yyyy") & "-" & Format(mReport.E12_ToPYCurr, "dd/MM/yyyy"))
            .MergeWorksheetCells(1, 9, 1, 10)
            .SetCellStyle(2, 9, xlNumStyle)
            .SetCellValue(2, 9, "Prev.Year Month")
            .MergeWorksheetCells(2, 9, 2, 10)
            .SetCellStyle(1, 11, xlNumStyle)
            .SetCellValue(1, 11, Format(mReport.E12_FromPYTD, "dd/MM/yyyy") & "-" & Format(mReport.E12_ToPYTD, "dd/MM/yyyy"))
            .MergeWorksheetCells(1, 11, 1, 12)
            .SetCellStyle(2, 11, xlNumStyle)
            .SetCellValue(2, 11, "Previous Year To Date")
            .MergeWorksheetCells(2, 11, 2, 12)
            .SetCellStyle(3, 1, 3, 12, xlNumStyle)
            .SetCellValue(3, 1, "Group Name")
            .SetCellValue(3, 2, "Client")
            .SetCellValue(3, 3, "Curr.Month")
            .SetCellValue(3, 4, "Budget")
            .SetCellValue(3, 5, "Comparison")
            .SetCellValue(3, 6, "YTD")
            .SetCellValue(3, 7, "Budget")
            .SetCellValue(3, 8, "Comparison")
            .SetCellValue(3, 9, "Prev.Year Month")
            .SetCellValue(3, 10, "Comparison")
            .SetCellValue(3, 11, "Prev.YTD")
            .SetCellValue(3, 12, "Comparison")
        End With
    End Sub
    Public Sub E13_TicketAnalysis(ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "TicketAnalysis")
                .FreezePanes(3, 0)
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(25, 31, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(32, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(12, xlNumStyle)
                .SetColumnStyle(16, xlNumStyle)
                .SetColumnStyle(21, xlNumStyle)
                xlNumStyle.FormatCode = "@"
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    .SetCellValue(3, j + 1, mdsDataSet.Tables(0).Columns(j).Caption)
                Next
            End With

            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    xlWorkSheet.SetCellValue(xlINVCount + 3, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                Next
            Next

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E15_DailyProfitReport(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        ' Data SS Description
        '	00	  Tots
        '	01	  ClientGroupDescription
        '	02 01 ClientCode
        '	03 02 ClientName
        'CURRENT AIR
        '	04 03 NetPayableAir
        '	05 04 NetBuyAIR
        '	06 05   IW5AIR
        '	07 06   IW6AIR
        '	08 07   IW7AIR
        '	09 08   IW8AIR
        '	10 09   IW9AIR
        '                     IW11AIR
        '   11 10   IW10AIR
        '	12 11 IWAIR
        '	13 12 ProfitAir
        '	14 13 PaxAIR
        '	15 14 ProfitPerPaxAir
        'CURRENT SERVICES
        '	16 15 NetPayableServices
        '	17 16 NetBuyServices
        '	18 17   IW5Services
        '	19 18   IW6Services
        '	20 19   IW7Services
        '	21 20   IW8Services
        '	22 21   IW9Services
        '                     IW11Services
        '	23 22   IW10Services
        '	24 23 IWServices
        '	25 24 ProfitServices
        '	26 25 PaxServices
        '	27 26 ProfitPerPaxServices
        'CURRENT TOTAL
        '	28 27 NetPayable
        '	29 28 NetBuy
        '	30 29   IW5
        '	31 30   IW6
        '	32 31   IW7
        '	33 32   IW8
        '	34 33   IW9
        '                     IW11
        '	35 34   IW10
        '	36 35	IW
        '	37 36 Profit
        '	38 37 Pax
        '	39 38 ProfitPerPax
        'YTD AIR
        '	40	  NetPayableYTDAir
        '	41	  NetBuyYTDAIR
        '	42	    IW5YTDAIR
        '	43	    IW6YTDAIR
        '	44	    IW7YTDAIR
        '	45	    IW8YTDAIR
        '	46	    IW9YTDAIR
        '                     IW11YTDAIR
        '	47	    IW10YTDAIR
        '	48	  IWYTDAIR
        '	49	  ProfitYTDAir
        '	50	  PaxYTDAIR
        '	51	  ProfitPerPaxYTDAir
        'YTD SERVICES
        '	52	  NetPayableYTDServices
        '	53	  NetBuyYTDServices
        '	54	    IW5YTDServices
        '	55	    IW6YTDServices
        '	56	    IW7YTDServices
        '	57	    IW8YTDServices
        '	58	    IW9YTDServices
        '                     IW11YTDServices
        '	59	    IW10YTDServices
        '	60	  IWYTDServices
        '	61	  ProfitYTDServices
        '	62	  PaxYTDServices
        '	63	  ProfitPerPaxYTDServices
        'YTD TOTAL
        '	64	  NetPayableYTD
        '	65	  NetBuyYTD
        '	66	    IW5YTD
        '	67	    IW6YTD
        '	68	    IW7YTD
        '	69	    IW8YTD
        '	70	    IW9YTD
        '                     IW11YTD
        '	71	    IW10YTD
        '	72	  IWYTD
        '	73	  ProfitYTD
        '	74	  PaxYTD
        '	75 39 ProfitPerPaxYTD

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 2)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 45, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(40, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(44, xlNumStyle)
                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 16, "Other Services")
                .SetCellValue(2, 29, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")

                For i = 3 To 29 Step 13
                    .SetCellValue(3, i, "Net Payable")
                    .SetCellValue(3, i + 1, "Net Buy")
                    .SetCellValue(3, i + 2, "IW5")
                    .SetCellValue(3, i + 3, "IW6")
                    .SetCellValue(3, i + 4, "IW7")
                    .SetCellValue(3, i + 5, "IW8")
                    .SetCellValue(3, i + 6, "IW9")
                    .SetCellValue(3, i + 7, "IW11")
                    .SetCellValue(3, i + 8, "IW10")
                    .SetCellValue(3, i + 9, "IW")
                    .SetCellValue(3, i + 10, "Profit")
                    .SetCellValue(3, i + 11, "Pax")
                    .SetCellValue(3, i + 12, "Profit Per Pax")
                Next

                .SetCellValue(2, 42, "YTD")
                .SetCellValue(3, 42, "Profit Per Pax")
                .SetCellValue(2, 43, "YTD")
                .SetCellValue(3, 43, "Pax")

                .SetCellValue(2, 44, "Provisional")
                .SetCellValue(3, 44, "Uninvoiced T/O")
                Dim pCommentTO As SpreadsheetLight.SLComment = .CreateComment
                pCommentTO.SetText("Approximate value. Markup or sales fee might still be outstanding")
                .InsertComment(2, 44, pCommentTO)

                .SetCellValue(3, 45, "Uninvoiced Pax")
                .SetCellValue(2, 46, "Provisional Profit")
                .SetCellValue(3, 46, "from uninvoiced Pax")
                Dim pComment As SpreadsheetLight.SLComment = .CreateComment
                pComment.SetText("Approximate value using weighted average")
                .InsertComment(2, 46, pComment)
                .SetCellStyle(1, 3, 3, 46, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 41)
                .MergeWorksheetCells(2, 3, 2, 15)
                .MergeWorksheetCells(2, 16, 2, 28)
                .MergeWorksheetCells(2, 29, 2, 41)


                xlINVCount = 3
                Dim pTotals(85) As Decimal
                Dim pOtherClients(85) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                Dim pProfitUninvoiced As Decimal = 0
                Dim pProfitUninvoicedRows(mdsDataSet.Tables(0).Rows.Count) As Decimal
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' row type 1 to be shown always
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" Then
                        ' if this is a client group get sum from details
                        If mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            pProfitUninvoicedRows(i) = 0
                            ' Read all rows to find type 2 for this client group
                            For ii = i + 1 To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    pProfitUninvoiced = mdsDataSet.Tables(0).Rows(ii).Item(81) * mdsDataSet.Tables(0).Rows(ii).Item(82)
                                    If pProfitUninvoiced > 0 Then
                                        pProfitUninvoicedRows(i) += pProfitUninvoiced
                                        pProfitUninvoicedRows(ii) = pProfitUninvoiced
                                    Else
                                        pProfitUninvoicedRows(ii) = 0
                                    End If
                                End If
                            Next
                        Else
                            ' this is an individual client
                            pProfitUninvoiced = mdsDataSet.Tables(0).Rows(i).Item(81) * mdsDataSet.Tables(0).Rows(i).Item(82)
                            If pProfitUninvoiced > 0 Then
                                pProfitUninvoicedRows(i) = pProfitUninvoiced
                            Else
                                pProfitUninvoicedRows(i) = 0
                            End If
                        End If
                    End If
                Next
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' Report Lines indicated by 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" And (pTopClients = 0 Or pTopClientsCount < pTopClients) Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                ' Detail rows
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 42
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    ' profit per pax YTD
                                    .SetCellValue(xlINVCount, 42, mdsDataSet.Tables(0).Rows(ii).Item(81))
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(42) < mdsDataSet.Tables(0).Rows(ii).Item(81) Then
                                        .SetCellStyle(xlINVCount, 41, xlStyleDetailNeg)
                                    End If
                                    ' pax YTD
                                    .SetCellValue(xlINVCount, 43, mdsDataSet.Tables(0).Rows(ii).Item(80))
                                    ' uninvoiced pax
                                    If mdsDataSet.Tables(0).Rows(ii).Item(82) <> 0 Or mdsDataSet.Tables(0).Rows(ii).Item(84) <> 0 Then
                                        .SetCellValue(xlINVCount, 44, mdsDataSet.Tables(0).Rows(ii).Item(84)) ' Net Payable uninvoiced
                                        .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(ii).Item(82)) ' Pax uninvoiced
                                        If pProfitUninvoicedRows(ii) <> 0 Then
                                            .SetCellValue(xlINVCount, 46, pProfitUninvoicedRows(ii))
                                        End If
                                    End If
                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        ' customer group or individual customers
                        xlINVCount += 1
                        For j = 2 To 42
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pTotals(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pTotals(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd
                        .SetCellValue(xlINVCount, 42, mdsDataSet.Tables(0).Rows(i).Item(81)) ' profit per pax ytd
                        If mdsDataSet.Tables(0).Rows(i).Item(42) < mdsDataSet.Tables(0).Rows(i).Item(81) Then
                            .SetCellStyle(xlINVCount, 41, xlStyleNegative)
                        End If
                        .SetCellValue(xlINVCount, 43, mdsDataSet.Tables(0).Rows(i).Item(80)) ' pax ytd
                        ' uninvoiced pax
                        If mdsDataSet.Tables(0).Rows(i).Item(82) <> 0 Or mdsDataSet.Tables(0).Rows(i).Item(84) <> 0 Then
                            .SetCellValue(xlINVCount, 44, mdsDataSet.Tables(0).Rows(i).Item(84)) ' Net Payable uninvoiced
                            .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(i).Item(82)) ' Pax uninvoiced
                            .SetCellValue(xlINVCount, 46, pProfitUninvoicedRows(i))
                            pTotals(82) += mdsDataSet.Tables(0).Rows(i).Item(82)
                            pTotals(84) += mdsDataSet.Tables(0).Rows(i).Item(84)
                            pTotals(85) += pProfitUninvoicedRows(i)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 42
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pOtherClients(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pOtherClients(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd
                        pOtherClients(82) += mdsDataSet.Tables(0).Rows(i).Item(82) ' pax uninvoiced
                        pOtherClients(84) += mdsDataSet.Tables(0).Rows(i).Item(84) ' net payable uninvoiced
                        pOtherClients(85) += pProfitUninvoicedRows(i)
                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(15) <> 0 Then
                        pOtherClients(16) = pOtherClients(14) / pOtherClients(15)
                    Else
                        pOtherClients(16) = 0
                    End If
                    If pOtherClients(28) <> 0 Then
                        pOtherClients(29) = pOtherClients(27) / pOtherClients(28)
                    Else
                        pOtherClients(29) = 0
                    End If
                    If pOtherClients(41) <> 0 Then
                        pOtherClients(42) = pOtherClients(40) / pOtherClients(41)
                    Else
                        pOtherClients(42) = 0
                    End If
                    If pOtherClients(80) <> 0 Then
                        pOtherClients(81) = pOtherClients(79) / pOtherClients(80)
                    Else
                        pOtherClients(81) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 42
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    pTotals(79) += pOtherClients(79) ' profit ytd
                    pTotals(80) += pOtherClients(80) ' pax ytd
                    .SetCellValue(xlINVCount, 42, pOtherClients(81))
                    If pOtherClients(42) < pOtherClients(81) Then
                        .SetCellStyle(xlINVCount, 41, xlStyleNegative)
                    End If
                    .SetCellValue(xlINVCount, 43, pOtherClients(80)) ' pax ytd
                    ' uninvoiced pax
                    If pOtherClients(82) <> 0 Or pOtherClients(84) <> 0 Then
                        .SetCellValue(xlINVCount, 44, pOtherClients(84))
                        .SetCellValue(xlINVCount, 45, pOtherClients(82))
                        .SetCellValue(xlINVCount, 46, pOtherClients(85))
                        pTotals(82) += pOtherClients(82) ' pax uninvoiced
                        pTotals(84) += pOtherClients(84) ' Net Payable uninvoiced
                        pTotals(85) += pOtherClients(85)
                    End If
                End If
                If pTotals(15) <> 0 Then
                    pTotals(16) = pTotals(14) / pTotals(15)
                Else
                    pTotals(16) = 0
                End If
                If pTotals(28) <> 0 Then
                    pTotals(29) = pTotals(27) / pTotals(28)
                Else
                    pTotals(29) = 0
                End If
                If pTotals(41) <> 0 Then
                    pTotals(42) = pTotals(40) / pTotals(41)
                Else
                    pTotals(42) = 0
                End If
                If pTotals(80) <> 0 Then
                    pTotals(81) = pTotals(79) / pTotals(80)
                Else
                    pTotals(81) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 42
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 42, pTotals(81))
                If pTotals(42) < pTotals(81) Then
                    .SetCellStyle(xlINVCount + 1, 41, xlStyleNegative)
                End If
                .SetCellValue(xlINVCount + 1, 43, pTotals(80))
                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 46, xlStyleHeader)
                ' uninvoiced pax
                If pTotals(82) <> 0 Or pTotals(84) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 44, pTotals(84))
                    .InsertComment(xlINVCount + 1, 44, pCommentTO)
                    .SetCellValue(xlINVCount + 1, 45, pTotals(82))
                    .SetCellValue(xlINVCount + 1, 46, pTotals(85))
                End If

                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 15, xlINVCount + 1, 15, xlStyleHeader)
                .SetCellStyle(3, 28, xlINVCount + 1, 28, xlStyleHeader)
                .SetCellStyle(3, 41, xlINVCount + 1, 42, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 42, xlNumStyle)
                .SetColumnStyle(3, 44, xlNumStyle)
                .SetColumnStyle(3, 46, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(40, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(45, xlNumStyle)

                .SetColumnStyle(5, 11, xlStyleDetail)
                .SetColumnStyle(18, 24, xlStyleDetail)
                .SetColumnStyle(31, 37, xlStyleDetail)

                .AutoFitColumn(0, 46)
                .GroupColumns(5, 11)
                .CollapseColumns(12)
                .GroupColumns(18, 24)
                .CollapseColumns(25)
                .GroupColumns(31, 37)
                .CollapseColumns(38)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Sub E41_DailyProfitReportWithRINVAAnalysis(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        ' Data SS Description
        '	00	  Tots
        '	01	  ClientGroupDescription
        '	02 01 ClientCode
        '	03 02 ClientName
        'CURRENT AIR
        '	04 03 NetPayableAir
        '	05 04 NetBuyAIR
        '	06 05   IW5AIR
        '	07 06   IW6AIR
        '	08 07   IW7AIR
        '	09 08   IW8AIR
        '	10 09   IW9AIR
        '                     IW11AIR
        '   11 10   IW10AIR
        '	12 11 IWAIR
        '	13 12 ProfitAir
        '	14 13 PaxAIR
        '	15 14 ProfitPerPaxAir
        'CURRENT SERVICES
        '	16 15 NetPayableServices
        '	17 16 NetBuyServices
        '	18 17   IW5Services
        '	19 18   IW6Services
        '	20 19   IW7Services
        '	21 20   IW8Services
        '	22 21   IW9Services
        '                     IW11Services
        '	23 22   IW10Services
        '	24 23 IWServices
        '	25 24 ProfitServices
        '	26 25 PaxServices
        '	27 26 ProfitPerPaxServices
        'CURRENT TOTAL
        '	28 27 NetPayable
        '	29 28 NetBuy
        '	30 29   IW5
        '	31 30   IW6
        '	32 31   IW7
        '	33 32   IW8
        '	34 33   IW9
        '                     IW11
        '	35 34   IW10
        '	36 35	IW
        '	37 36 Profit
        '	38 37 Pax
        '	39 38 ProfitPerPax
        'YTD AIR
        '	40	  NetPayableYTDAir
        '	41	  NetBuyYTDAIR
        '	42	    IW5YTDAIR
        '	43	    IW6YTDAIR
        '	44	    IW7YTDAIR
        '	45	    IW8YTDAIR
        '	46	    IW9YTDAIR
        '                     IW11YTDAIR
        '	47	    IW10YTDAIR
        '	48	  IWYTDAIR
        '	49	  ProfitYTDAir
        '	50	  PaxYTDAIR
        '	51	  ProfitPerPaxYTDAir
        'YTD SERVICES
        '	52	  NetPayableYTDServices
        '	53	  NetBuyYTDServices
        '	54	    IW5YTDServices
        '	55	    IW6YTDServices
        '	56	    IW7YTDServices
        '	57	    IW8YTDServices
        '	58	    IW9YTDServices
        '                     IW11YTDServices
        '	59	    IW10YTDServices
        '	60	  IWYTDServices
        '	61	  ProfitYTDServices
        '	62	  PaxYTDServices
        '	63	  ProfitPerPaxYTDServices
        'YTD TOTAL
        '	64	  NetPayableYTD
        '	65	  NetBuyYTD
        '	66	    IW5YTD
        '	67	    IW6YTD
        '	68	    IW7YTD
        '	69	    IW8YTD
        '	70	    IW9YTD
        '                     IW11YTD
        '	71	    IW10YTD
        '	72	  IWYTD
        '	73	  ProfitYTD
        '	74	  PaxYTD
        '	75 39 ProfitPerPaxYTD

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 2)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleRINVA As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRINVA.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleRINVA.Fill.SetPatternForegroundColor(Color.LightGreen)
                xlStyleRINVA.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 48, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(46, xlNumStyle)
                .SetColumnStyle(47, xlNumStyle)
                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 16, "Other Services")
                .SetCellValue(2, 29, "RINVA")
                .SetCellValue(2, 32, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")

                Dim i As Integer
                For i1 = 1 To 3
                    i = 3
                    If i1 = 2 Then i = 16
                    If i1 = 3 Then i = 32

                    .SetCellValue(3, i, "Net Payable")
                    .SetCellValue(3, i + 1, "Net Buy")
                    .SetCellValue(3, i + 2, "IW5")
                    .SetCellValue(3, i + 3, "IW6")
                    .SetCellValue(3, i + 4, "IW7")
                    .SetCellValue(3, i + 5, "IW8")
                    .SetCellValue(3, i + 6, "IW9")
                    .SetCellValue(3, i + 7, "IW11")
                    .SetCellValue(3, i + 8, "IW10")
                    .SetCellValue(3, i + 9, "IW")
                    .SetCellValue(3, i + 10, "Profit")
                    .SetCellValue(3, i + 11, "Pax")
                    .SetCellValue(3, i + 12, "Profit Per Pax")
                Next
                .SetCellValue(3, 29, "Net Payable")
                .SetCellValue(3, 30, "Net Buy")
                .SetCellValue(3, 31, "Profit")

                .SetCellValue(2, 45, "YTD")
                .SetCellValue(3, 45, "Profit Per Pax")
                .SetCellValue(2, 46, "YTD")
                .SetCellValue(3, 47, "Pax")

                .SetCellValue(2, 47, "Provisional")
                .SetCellValue(3, 47, "Uninvoiced T/O")
                Dim pCommentTO As SpreadsheetLight.SLComment = .CreateComment
                pCommentTO.SetText("Approximate value. Markup or sales fee might still be outstanding")
                .InsertComment(2, 47, pCommentTO)

                .SetCellValue(3, 48, "Uninvoiced Pax")
                .SetCellValue(2, 49, "Provisional Profit")
                .SetCellValue(3, 49, "from uninvoiced Pax")
                Dim pComment As SpreadsheetLight.SLComment = .CreateComment
                pComment.SetText("Approximate value using weighted average")
                .InsertComment(2, 49, pComment)
                .SetCellStyle(1, 3, 3, 49, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 44)
                .MergeWorksheetCells(2, 3, 2, 15)
                .MergeWorksheetCells(2, 16, 2, 28)
                .MergeWorksheetCells(2, 29, 2, 31)
                .MergeWorksheetCells(2, 32, 2, 44)


                xlINVCount = 3
                Dim pTotals(91) As Decimal
                Dim pOtherClients(91) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                Dim pProfitUninvoiced As Decimal = 0
                Dim pProfitUninvoicedRows(mdsDataSet.Tables(0).Rows.Count) As Decimal
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' row type 1 to be shown always
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" Then
                        ' if this is a client group get sum from details
                        If mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            pProfitUninvoicedRows(i) = 0
                            ' Read all rows to find type 2 for this client group
                            For ii = i + 1 To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    pProfitUninvoiced = mdsDataSet.Tables(0).Rows(ii).Item(81) * mdsDataSet.Tables(0).Rows(ii).Item(82)
                                    If pProfitUninvoiced > 0 Then
                                        pProfitUninvoicedRows(i) += pProfitUninvoiced
                                        pProfitUninvoicedRows(ii) = pProfitUninvoiced
                                    Else
                                        pProfitUninvoicedRows(ii) = 0
                                    End If
                                End If
                            Next
                        Else
                            ' this is an individual client
                            pProfitUninvoiced = mdsDataSet.Tables(0).Rows(i).Item(81) * mdsDataSet.Tables(0).Rows(i).Item(82)
                            If pProfitUninvoiced > 0 Then
                                pProfitUninvoicedRows(i) = pProfitUninvoiced
                            Else
                                pProfitUninvoicedRows(i) = 0
                            End If
                        End If
                    End If
                Next
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' Report Lines indicated by 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" And (pTopClients = 0 Or pTopClientsCount < pTopClients) Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                ' Detail rows
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 29
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    ' RINVA
                                    .SetCellValue(xlINVCount, 29, mdsDataSet.Tables(0).Rows(ii).Item(85))
                                    .SetCellValue(xlINVCount, 30, mdsDataSet.Tables(0).Rows(ii).Item(86))
                                    .SetCellValue(xlINVCount, 31, mdsDataSet.Tables(0).Rows(ii).Item(87))
                                    For jj As Integer = 30 To 42
                                        .SetCellValue(xlINVCount, jj + 2, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next

                                    ' profit per pax YTD
                                    .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(ii).Item(81))
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 49, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(42) < mdsDataSet.Tables(0).Rows(ii).Item(81) Then
                                        .SetCellStyle(xlINVCount, 44, xlStyleDetailNeg)
                                    End If
                                    ' pax YTD
                                    .SetCellValue(xlINVCount, 46, mdsDataSet.Tables(0).Rows(ii).Item(80))
                                    ' uninvoiced pax
                                    If mdsDataSet.Tables(0).Rows(ii).Item(82) <> 0 Or mdsDataSet.Tables(0).Rows(ii).Item(84) <> 0 Then
                                        .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(ii).Item(84)) ' Net Payable uninvoiced
                                        .SetCellValue(xlINVCount, 48, mdsDataSet.Tables(0).Rows(ii).Item(82)) ' Pax uninvoiced
                                        If pProfitUninvoicedRows(ii) <> 0 Then
                                            .SetCellValue(xlINVCount, 49, pProfitUninvoicedRows(ii))
                                        End If
                                    End If
                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        ' customer group or individual customers
                        xlINVCount += 1
                        For j = 2 To 29
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        ' RINVA
                        .SetCellValue(xlINVCount, 29, mdsDataSet.Tables(0).Rows(i).Item(85))
                        .SetCellValue(xlINVCount, 30, mdsDataSet.Tables(0).Rows(i).Item(86))
                        .SetCellValue(xlINVCount, 31, mdsDataSet.Tables(0).Rows(i).Item(87))
                        pTotals(86) += mdsDataSet.Tables(0).Rows(i).Item(85)
                        pTotals(87) += mdsDataSet.Tables(0).Rows(i).Item(86)
                        pTotals(88) += mdsDataSet.Tables(0).Rows(i).Item(87)

                        For j = 30 To 42
                            .SetCellValue(xlINVCount, j + 2, mdsDataSet.Tables(0).Rows(i).Item(j))
                            pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        Next
                        pTotals(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pTotals(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd
                        .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(i).Item(81)) ' profit per pax ytd
                        If mdsDataSet.Tables(0).Rows(i).Item(42) < mdsDataSet.Tables(0).Rows(i).Item(81) Then
                            .SetCellStyle(xlINVCount, 44, xlStyleNegative)
                        End If
                        .SetCellValue(xlINVCount, 46, mdsDataSet.Tables(0).Rows(i).Item(80)) ' pax ytd
                        ' uninvoiced pax
                        If mdsDataSet.Tables(0).Rows(i).Item(82) <> 0 Or mdsDataSet.Tables(0).Rows(i).Item(84) <> 0 Then
                            .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(i).Item(84)) ' Net Payable uninvoiced
                            .SetCellValue(xlINVCount, 48, mdsDataSet.Tables(0).Rows(i).Item(82)) ' Pax uninvoiced
                            .SetCellValue(xlINVCount, 49, pProfitUninvoicedRows(i))
                            pTotals(82) += mdsDataSet.Tables(0).Rows(i).Item(82)
                            pTotals(84) += mdsDataSet.Tables(0).Rows(i).Item(84)
                            pTotals(85) += pProfitUninvoicedRows(i)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 42
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pOtherClients(86) += mdsDataSet.Tables(0).Rows(i).Item(85) ' RINVA Payable
                        pOtherClients(87) += mdsDataSet.Tables(0).Rows(i).Item(86) ' RINVA Buy
                        pOtherClients(88) += mdsDataSet.Tables(0).Rows(i).Item(87) ' RINVA Profit

                        pOtherClients(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pOtherClients(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd
                        pOtherClients(82) += mdsDataSet.Tables(0).Rows(i).Item(82) ' pax uninvoiced
                        pOtherClients(84) += mdsDataSet.Tables(0).Rows(i).Item(84) ' net payable uninvoiced
                        pOtherClients(85) += pProfitUninvoicedRows(i)
                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(15) <> 0 Then
                        pOtherClients(16) = pOtherClients(14) / pOtherClients(15)
                    Else
                        pOtherClients(16) = 0
                    End If
                    If pOtherClients(28) <> 0 Then
                        pOtherClients(29) = pOtherClients(27) / pOtherClients(28)
                    Else
                        pOtherClients(29) = 0
                    End If
                    If pOtherClients(41) <> 0 Then
                        pOtherClients(42) = pOtherClients(40) / pOtherClients(41)
                    Else
                        pOtherClients(42) = 0
                    End If
                    If pOtherClients(80) <> 0 Then
                        pOtherClients(81) = pOtherClients(79) / pOtherClients(80)
                    Else
                        pOtherClients(81) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 29
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    .SetCellValue(xlINVCount, 29, pOtherClients(86))
                    .SetCellValue(xlINVCount, 30, pOtherClients(87))
                    .SetCellValue(xlINVCount, 31, pOtherClients(88))
                    For j = 30 To 42
                        .SetCellValue(xlINVCount, j + 2, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next

                    pTotals(79) += pOtherClients(79) ' profit ytd
                    pTotals(80) += pOtherClients(80) ' pax ytd
                    .SetCellValue(xlINVCount, 45, pOtherClients(81))
                    If pOtherClients(42) < pOtherClients(81) Then
                        .SetCellStyle(xlINVCount, 44, xlStyleNegative)
                    End If
                    .SetCellValue(xlINVCount, 46, pOtherClients(80)) ' pax ytd
                    ' uninvoiced pax
                    If pOtherClients(82) <> 0 Or pOtherClients(84) <> 0 Then
                        .SetCellValue(xlINVCount, 47, pOtherClients(84))
                        .SetCellValue(xlINVCount, 48, pOtherClients(82))
                        .SetCellValue(xlINVCount, 49, pOtherClients(85))
                        pTotals(82) += pOtherClients(82) ' pax uninvoiced
                        pTotals(84) += pOtherClients(84) ' Net Payable uninvoiced
                        pTotals(85) += pOtherClients(85)
                    End If
                End If
                If pTotals(15) <> 0 Then
                    pTotals(16) = pTotals(14) / pTotals(15)
                Else
                    pTotals(16) = 0
                End If
                If pTotals(28) <> 0 Then
                    pTotals(29) = pTotals(27) / pTotals(28)
                Else
                    pTotals(29) = 0
                End If
                If pTotals(41) <> 0 Then
                    pTotals(42) = pTotals(40) / pTotals(41)
                Else
                    pTotals(42) = 0
                End If
                If pTotals(80) <> 0 Then
                    pTotals(81) = pTotals(79) / pTotals(80)
                Else
                    pTotals(81) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 29
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 29, pTotals(86))
                .SetCellValue(xlINVCount + 1, 30, pTotals(87))
                .SetCellValue(xlINVCount + 1, 31, pTotals(88))
                For j = 30 To 42
                    .SetCellValue(xlINVCount + 1, j + 2, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 45, pTotals(81))
                If pTotals(42) < pTotals(81) Then
                    .SetCellStyle(xlINVCount + 1, 44, xlStyleNegative)
                End If
                .SetCellValue(xlINVCount + 1, 46, pTotals(80))
                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 49, xlStyleHeader)
                ' uninvoiced pax
                If pTotals(82) <> 0 Or pTotals(84) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 47, pTotals(84))
                    .InsertComment(xlINVCount + 1, 47, pCommentTO)
                    .SetCellValue(xlINVCount + 1, 48, pTotals(82))
                    .SetCellValue(xlINVCount + 1, 49, pTotals(85))
                End If

                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 15, xlINVCount + 1, 15, xlStyleHeader)
                .SetCellStyle(3, 28, xlINVCount + 1, 28, xlStyleHeader)
                .SetCellStyle(3, 29, xlINVCount + 1, 31, xlStyleRINVA)
                .SetCellStyle(3, 44, xlINVCount + 1, 45, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 45, xlNumStyle)
                .SetColumnStyle(3, 47, xlNumStyle)
                .SetColumnStyle(3, 49, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(46, xlNumStyle)
                .SetColumnStyle(48, xlNumStyle)

                .SetColumnStyle(5, 11, xlStyleDetail)
                .SetColumnStyle(18, 24, xlStyleDetail)
                .SetColumnStyle(34, 40, xlStyleDetail)

                .AutoFitColumn(0, 49)
                .GroupColumns(5, 11)
                .CollapseColumns(12)
                .GroupColumns(18, 24)
                .CollapseColumns(25)
                .GroupColumns(34, 40)
                .CollapseColumns(41)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E43_DailyProfitReportWithProvisionalAnalysis(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 2)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleRINVA As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRINVA.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleRINVA.Fill.SetPatternForegroundColor(Color.LightGreen)
                xlStyleRINVA.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 53, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(46, xlNumStyle)
                .SetColumnStyle(48, xlNumStyle)
                .SetColumnStyle(52, xlNumStyle)

                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(50, xlNumStyle)
                .SetColumnStyle(54, xlNumStyle)

                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                If mReport.Date2Checked Then
                    .SetCellValue(1, 47, $"Provisional to {mReport.Date2From:dd/MM/yyyy}")
                    .SetCellValue(1, 51, "Provisional later")
                Else
                    .SetCellValue(1, 47, "Provisional")
                End If
                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 16, "Other Services")
                .SetCellValue(2, 29, "RINVA")
                .SetCellValue(2, 32, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")

                Dim i As Integer
                For i1 = 1 To 3
                    i = 3
                    If i1 = 2 Then i = 16
                    If i1 = 3 Then i = 32

                    .SetCellValue(3, i, "Net Payable")
                    .SetCellValue(3, i + 1, "Net Buy")
                    .SetCellValue(3, i + 2, "IW5")
                    .SetCellValue(3, i + 3, "IW6")
                    .SetCellValue(3, i + 4, "IW7")
                    .SetCellValue(3, i + 5, "IW8")
                    .SetCellValue(3, i + 6, "IW9")
                    .SetCellValue(3, i + 7, "IW11")
                    .SetCellValue(3, i + 8, "IW10")
                    .SetCellValue(3, i + 9, "IW")
                    .SetCellValue(3, i + 10, "Profit")
                    .SetCellValue(3, i + 11, "Pax")
                    .SetCellValue(3, i + 12, "Profit Per Pax")
                Next
                .SetCellValue(3, 29, "Net Payable")
                .SetCellValue(3, 30, "Net Buy")
                .SetCellValue(3, 31, "Profit")

                .SetCellValue(2, 45, "YTD")
                .SetCellValue(3, 45, "Profit Per Pax")
                .SetCellValue(2, 46, "YTD")
                .SetCellValue(3, 46, "Pax")

                .SetCellValue(2, 47, "Provisional")
                .SetCellValue(3, 47, "Uninvoiced T/O")
                Dim pCommentTO As SpreadsheetLight.SLComment = .CreateComment
                pCommentTO.SetText("Approximate value. Markup or sales fee might still be outstanding")
                .InsertComment(2, 47, pCommentTO)

                .SetCellValue(3, 48, "Uninvoiced Pax")
                .SetCellValue(2, 49, "Provisional Profit")
                .SetCellValue(3, 49, "from uninvoiced Pax")
                Dim pComment As SpreadsheetLight.SLComment = .CreateComment
                pComment.SetText("Approximate value using weighted average")
                .InsertComment(2, 49, pComment)
                .SetCellValue(3, 50, "To Date")

                If mReport.Date2Checked Then
                    .SetCellValue(2, 51, "Provisional")
                    .SetCellValue(3, 51, "Uninvoiced T/O")
                    .SetCellValue(3, 52, "Uninvoiced Pax")
                    .SetCellValue(2, 53, "Provisional Profit")
                    .SetCellValue(3, 53, "from uninvoiced Pax")
                    .SetCellValue(3, 54, "From Date")
                End If
                .SetCellStyle(1, 3, 3, 54, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 44)
                .MergeWorksheetCells(2, 3, 2, 15)
                .MergeWorksheetCells(2, 16, 2, 28)
                .MergeWorksheetCells(2, 29, 2, 31)
                .MergeWorksheetCells(2, 32, 2, 44)

                .MergeWorksheetCells(1, 47, 1, 50)
                .MergeWorksheetCells(1, 51, 1, 54)

                xlINVCount = 3
                Dim pTotals(94) As Decimal
                Dim pOtherClients(94) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                Dim pProfitUninvoiced As Decimal = 0
                Dim pProfitUninvoicedRows(mdsDataSet.Tables(0).Rows.Count) As Decimal
                Dim pProfitUninvoicedLater As Decimal = 0
                Dim pProfitUninvoicedRowsLater(mdsDataSet.Tables(0).Rows.Count) As Decimal
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' row type 1 to be shown always
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" Then
                        ' if this is a client group get sum from details
                        If mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            pProfitUninvoicedRows(i) = 0
                            pProfitUninvoicedRowsLater(i) = 0

                            ' Read all rows to find type 2 for this client group
                            For ii = i + 1 To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    pProfitUninvoiced = mdsDataSet.Tables(0).Rows(ii).Item(81) * mdsDataSet.Tables(0).Rows(ii).Item(82)
                                    If pProfitUninvoiced > 0 Then
                                        pProfitUninvoicedRows(i) += pProfitUninvoiced
                                        pProfitUninvoicedRows(ii) = pProfitUninvoiced
                                    Else
                                        pProfitUninvoicedRows(ii) = 0
                                    End If
                                    pProfitUninvoicedLater = mdsDataSet.Tables(0).Rows(ii).Item(81) * mdsDataSet.Tables(0).Rows(ii).Item(91)
                                    If pProfitUninvoicedLater > 0 Then
                                        pProfitUninvoicedRowsLater(i) += pProfitUninvoicedLater
                                        pProfitUninvoicedRowsLater(ii) = pProfitUninvoicedLater
                                    Else
                                        pProfitUninvoicedRowsLater(ii) = 0
                                    End If
                                End If
                            Next
                        Else
                            ' this is an individual client
                            pProfitUninvoiced = mdsDataSet.Tables(0).Rows(i).Item(81) * mdsDataSet.Tables(0).Rows(i).Item(82)
                            If pProfitUninvoiced > 0 Then
                                pProfitUninvoicedRows(i) = pProfitUninvoiced
                            Else
                                pProfitUninvoicedRows(i) = 0
                            End If

                            pProfitUninvoicedLater = mdsDataSet.Tables(0).Rows(i).Item(81) * mdsDataSet.Tables(0).Rows(i).Item(91)
                            If pProfitUninvoicedLater > 0 Then
                                pProfitUninvoicedRowsLater(i) = pProfitUninvoicedLater
                            Else
                                pProfitUninvoicedRowsLater(i) = 0
                            End If
                        End If
                    End If
                Next
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' Report Lines indicated by 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" And (pTopClients = 0 Or pTopClientsCount < pTopClients) Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                ' Detail rows
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 29
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    ' RINVA
                                    .SetCellValue(xlINVCount, 29, mdsDataSet.Tables(0).Rows(ii).Item(85))
                                    .SetCellValue(xlINVCount, 30, mdsDataSet.Tables(0).Rows(ii).Item(86))
                                    .SetCellValue(xlINVCount, 31, mdsDataSet.Tables(0).Rows(ii).Item(87))
                                    For jj As Integer = 30 To 42
                                        .SetCellValue(xlINVCount, jj + 2, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next

                                    ' profit per pax YTD
                                    .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(ii).Item(81))
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 54, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(42) < mdsDataSet.Tables(0).Rows(ii).Item(81) Then
                                        .SetCellStyle(xlINVCount, 44, xlStyleDetailNeg)
                                    End If
                                    ' pax YTD
                                    .SetCellValue(xlINVCount, 46, mdsDataSet.Tables(0).Rows(ii).Item(80))
                                    ' uninvoiced pax
                                    If mdsDataSet.Tables(0).Rows(ii).Item(82) <> 0 Or Not IsDBNull(mdsDataSet.Tables(0).Rows(ii).Item(94)) Then
                                        .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(ii).Item(84)) ' Net Payable uninvoiced
                                        .SetCellValue(xlINVCount, 48, mdsDataSet.Tables(0).Rows(ii).Item(82)) ' Pax uninvoiced
                                        If pProfitUninvoicedRows(ii) <> 0 Then
                                            .SetCellValue(xlINVCount, 49, pProfitUninvoicedRows(ii))
                                        End If
                                        .SetCellValue(xlINVCount, 50, mdsDataSet.Tables(0).Rows(ii).Item(94)) ' Max Dep Date
                                    End If
                                    ' uninvoiced later pax
                                    If mdsDataSet.Tables(0).Rows(ii).Item(91) <> 0 Or Not IsDBNull(mdsDataSet.Tables(0).Rows(ii).Item(95)) Then
                                        .SetCellValue(xlINVCount, 51, mdsDataSet.Tables(0).Rows(ii).Item(93)) ' Net Payable uninvoiced later
                                        .SetCellValue(xlINVCount, 52, mdsDataSet.Tables(0).Rows(ii).Item(91)) ' Pax uninvoiced later
                                        If pProfitUninvoicedRows(ii) <> 0 Then
                                            .SetCellValue(xlINVCount, 53, pProfitUninvoicedRowsLater(ii))
                                        End If
                                        .SetCellValue(xlINVCount, 54, mdsDataSet.Tables(0).Rows(ii).Item(95)) ' Min Dep Date
                                    End If

                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        ' customer group or individual customers
                        xlINVCount += 1
                        For j = 2 To 29
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        ' RINVA
                        .SetCellValue(xlINVCount, 29, mdsDataSet.Tables(0).Rows(i).Item(85))
                        .SetCellValue(xlINVCount, 30, mdsDataSet.Tables(0).Rows(i).Item(86))
                        .SetCellValue(xlINVCount, 31, mdsDataSet.Tables(0).Rows(i).Item(87))
                        pTotals(86) += mdsDataSet.Tables(0).Rows(i).Item(85)
                        pTotals(87) += mdsDataSet.Tables(0).Rows(i).Item(86)
                        pTotals(88) += mdsDataSet.Tables(0).Rows(i).Item(87)

                        For j = 30 To 42
                            .SetCellValue(xlINVCount, j + 2, mdsDataSet.Tables(0).Rows(i).Item(j))
                            pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        Next
                        pTotals(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pTotals(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd
                        .SetCellValue(xlINVCount, 45, mdsDataSet.Tables(0).Rows(i).Item(81)) ' profit per pax ytd
                        If mdsDataSet.Tables(0).Rows(i).Item(42) < mdsDataSet.Tables(0).Rows(i).Item(81) Then
                            .SetCellStyle(xlINVCount, 44, xlStyleNegative)
                        End If
                        .SetCellValue(xlINVCount, 46, mdsDataSet.Tables(0).Rows(i).Item(80)) ' pax ytd
                        ' uninvoiced pax
                        If mdsDataSet.Tables(0).Rows(i).Item(82) <> 0 Or Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(94)) Then
                            .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(i).Item(84)) ' Net Payable uninvoiced
                            .SetCellValue(xlINVCount, 48, mdsDataSet.Tables(0).Rows(i).Item(82)) ' Pax uninvoiced
                            .SetCellValue(xlINVCount, 49, pProfitUninvoicedRows(i))
                            If Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(94)) Then
                                .SetCellValue(xlINVCount, 50, mdsDataSet.Tables(0).Rows(i).Item(94))
                            End If
                            pTotals(82) += mdsDataSet.Tables(0).Rows(i).Item(82)
                            pTotals(84) += mdsDataSet.Tables(0).Rows(i).Item(84)
                            pTotals(85) += pProfitUninvoicedRows(i)
                        End If

                        ' uninvoiced pax later
                        If mdsDataSet.Tables(0).Rows(i).Item(91) <> 0 Or Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(95)) Then
                            .SetCellValue(xlINVCount, 51, mdsDataSet.Tables(0).Rows(i).Item(93)) ' Net Payable uninvoiced later
                            .SetCellValue(xlINVCount, 52, mdsDataSet.Tables(0).Rows(i).Item(91)) ' Pax uninvoiced later
                            .SetCellValue(xlINVCount, 53, pProfitUninvoicedRowsLater(i))
                            If Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(95)) Then
                                .SetCellValue(xlINVCount, 54, mdsDataSet.Tables(0).Rows(i).Item(95))
                            End If
                            pTotals(92) += mdsDataSet.Tables(0).Rows(i).Item(93)
                            pTotals(93) += mdsDataSet.Tables(0).Rows(i).Item(91)
                            pTotals(94) += pProfitUninvoicedRowsLater(i)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 42
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pOtherClients(86) += mdsDataSet.Tables(0).Rows(i).Item(85) ' RINVA Payable
                        pOtherClients(87) += mdsDataSet.Tables(0).Rows(i).Item(86) ' RINVA Buy
                        pOtherClients(88) += mdsDataSet.Tables(0).Rows(i).Item(87) ' RINVA Profit

                        pOtherClients(79) += mdsDataSet.Tables(0).Rows(i).Item(79) ' profit ytd
                        pOtherClients(80) += mdsDataSet.Tables(0).Rows(i).Item(80) ' pax ytd

                        pOtherClients(82) += mdsDataSet.Tables(0).Rows(i).Item(82) ' pax uninvoiced
                        pOtherClients(84) += mdsDataSet.Tables(0).Rows(i).Item(84) ' net payable uninvoiced
                        pOtherClients(85) += pProfitUninvoicedRows(i)

                        pOtherClients(92) += mdsDataSet.Tables(0).Rows(i).Item(93)
                        pOtherClients(93) += mdsDataSet.Tables(0).Rows(i).Item(91)
                        pOtherClients(94) += pProfitUninvoicedRowsLater(i)

                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(15) <> 0 Then
                        pOtherClients(16) = pOtherClients(14) / pOtherClients(15)
                    Else
                        pOtherClients(16) = 0
                    End If
                    If pOtherClients(28) <> 0 Then
                        pOtherClients(29) = pOtherClients(27) / pOtherClients(28)
                    Else
                        pOtherClients(29) = 0
                    End If
                    If pOtherClients(41) <> 0 Then
                        pOtherClients(42) = pOtherClients(40) / pOtherClients(41)
                    Else
                        pOtherClients(42) = 0
                    End If
                    If pOtherClients(80) <> 0 Then
                        pOtherClients(81) = pOtherClients(79) / pOtherClients(80)
                    Else
                        pOtherClients(81) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 29
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    .SetCellValue(xlINVCount, 29, pOtherClients(86))
                    .SetCellValue(xlINVCount, 30, pOtherClients(87))
                    .SetCellValue(xlINVCount, 31, pOtherClients(88))
                    For j = 30 To 42
                        .SetCellValue(xlINVCount, j + 2, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next

                    pTotals(79) += pOtherClients(79) ' profit ytd
                    pTotals(80) += pOtherClients(80) ' pax ytd
                    .SetCellValue(xlINVCount, 45, pOtherClients(81))
                    If pOtherClients(42) < pOtherClients(81) Then
                        .SetCellStyle(xlINVCount, 44, xlStyleNegative)
                    End If
                    .SetCellValue(xlINVCount, 46, pOtherClients(80)) ' pax ytd
                    ' uninvoiced pax
                    If pOtherClients(82) <> 0 Or pOtherClients(84) <> 0 Then
                        .SetCellValue(xlINVCount, 47, pOtherClients(84))
                        .SetCellValue(xlINVCount, 48, pOtherClients(82))
                        .SetCellValue(xlINVCount, 49, pOtherClients(85))
                        pTotals(82) += pOtherClients(82) ' pax uninvoiced
                        pTotals(84) += pOtherClients(84) ' Net Payable uninvoiced
                        pTotals(85) += pOtherClients(85)
                    End If

                    ' uninvoiced pax later
                    If pOtherClients(92) <> 0 Or pOtherClients(93) <> 0 Then
                        .SetCellValue(xlINVCount, 51, pOtherClients(93))
                        .SetCellValue(xlINVCount, 52, pOtherClients(92))
                        .SetCellValue(xlINVCount, 53, pOtherClients(94))
                        pTotals(92) += pOtherClients(92) ' pax uninvoiced
                        pTotals(93) += pOtherClients(93) ' Net Payable uninvoiced
                        pTotals(94) += pOtherClients(94)
                    End If

                End If
                If pTotals(15) <> 0 Then
                    pTotals(16) = pTotals(14) / pTotals(15)
                Else
                    pTotals(16) = 0
                End If
                If pTotals(28) <> 0 Then
                    pTotals(29) = pTotals(27) / pTotals(28)
                Else
                    pTotals(29) = 0
                End If
                If pTotals(41) <> 0 Then
                    pTotals(42) = pTotals(40) / pTotals(41)
                Else
                    pTotals(42) = 0
                End If
                If pTotals(80) <> 0 Then
                    pTotals(81) = pTotals(79) / pTotals(80)
                Else
                    pTotals(81) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 29
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 29, pTotals(86))
                .SetCellValue(xlINVCount + 1, 30, pTotals(87))
                .SetCellValue(xlINVCount + 1, 31, pTotals(88))
                For j = 30 To 42
                    .SetCellValue(xlINVCount + 1, j + 2, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 45, pTotals(81))
                If pTotals(42) < pTotals(81) Then
                    .SetCellStyle(xlINVCount + 1, 44, xlStyleNegative)
                End If
                .SetCellValue(xlINVCount + 1, 46, pTotals(80))
                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 54, xlStyleHeader)
                ' uninvoiced pax
                If pTotals(82) <> 0 Or pTotals(84) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 47, pTotals(84))
                    .InsertComment(xlINVCount + 1, 47, pCommentTO)
                    .SetCellValue(xlINVCount + 1, 48, pTotals(82))
                    .SetCellValue(xlINVCount + 1, 49, pTotals(85))
                End If
                ' uninvoiced pax later
                If pTotals(92) <> 0 Or pTotals(93) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 51, pTotals(92))
                    .SetCellValue(xlINVCount + 1, 52, pTotals(93))
                    .SetCellValue(xlINVCount + 1, 53, pTotals(94))
                End If

                xlStyleHeader.Fill.SetPatternForegroundColor(Color.LightGray)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                .SetCellStyle(3, 51, xlINVCount, 54, xlStyleHeader)

                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 15, xlINVCount + 1, 15, xlStyleHeader)
                .SetCellStyle(3, 28, xlINVCount + 1, 28, xlStyleHeader)
                .SetCellStyle(3, 29, xlINVCount + 1, 31, xlStyleRINVA)
                .SetCellStyle(3, 44, xlINVCount + 1, 45, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 53, xlNumStyle)

                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(14, xlNumStyle)
                .SetColumnStyle(27, xlNumStyle)
                .SetColumnStyle(43, xlNumStyle)
                .SetColumnStyle(46, xlNumStyle)
                .SetColumnStyle(48, xlNumStyle)
                .SetColumnStyle(52, xlNumStyle)
                xlNumStyle.FormatCode = "dd/MM/yyyy"
                .SetColumnStyle(50, xlNumStyle)
                .SetColumnStyle(54, xlNumStyle)
                .SetColumnStyle(5, 11, xlStyleDetail)
                .SetColumnStyle(18, 24, xlStyleDetail)
                .SetColumnStyle(34, 40, xlStyleDetail)

                .AutoFitColumn(0, 54)
                .GroupColumns(5, 11)
                .CollapseColumns(12)
                .GroupColumns(18, 24)
                .CollapseColumns(25)
                .GroupColumns(34, 40)
                .CollapseColumns(41)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E45_AirTicketSalesAll(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleGreyOut As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGreyOut.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGreyOut.Fill.SetPatternForegroundColor(Color.LightGray)

                Dim xlStyleMandatory As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleMandatory.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleMandatory.Fill.SetPatternForegroundColor(Color.Yellow)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 130, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(17, xlNumStyle)
                .SetColumnStyle(36, 38, xlNumStyle)
                .SetColumnStyle(133, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, xlNumStyle)
                .SetColumnStyle(15, xlNumStyle)
                .SetColumnStyle(28, 29, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(9, xlNumStyle)
                .SetColumnStyle(39, xlNumStyle)
                xlNumStyle.FormatCode = "HH:mm"
                .SetColumnStyle(132, xlNumStyle)

                .SetCellValue(1, 1, "Issue Date")
                .SetCellValue(1, 2, "Client Code")
                .SetCellValue(1, 3, "Client Name")
                .SetCellValue(1, 4, "Omit")
                .SetCellValue(1, 5, "Void")
                .SetCellValue(1, 6, "PNR")
                .SetCellValue(1, 7, "Ticket Number")
                .SetCellValue(1, 8, "Passenger")
                .SetCellValue(1, 9, "Pax Count")
                .SetCellValue(1, 10, "Product Type")
                .SetCellValue(1, 11, "Action Type")
                .SetCellValue(1, 12, "Inv Code")
                .SetCellValue(1, 13, "Inv Series")
                .SetCellValue(1, 14, "Inv Number")
                .SetCellValue(1, 15, "Invoice Date")
                .SetCellValue(1, 16, "Vessel")
                .SetCellValue(1, 17, "Net Payable")
                .SetCellValue(1, 18, "Verified")
                .SetCellValue(1, 19, "Remarks")
                .SetCellValue(1, 20, "Transaction Type")
                .SetCellValue(1, 21, "RegNr")
                .SetCellValue(1, 22, "Ticketing Airline")
                .SetCellValue(1, 23, "Routing")
                .SetCellValue(1, 24, "SalesPerson")
                .SetCellValue(1, 25, "Issuing Agent")
                .SetCellValue(1, 26, "Creator Agent")
                .SetCellValue(1, 27, "Reference")
                .SetCellValue(1, 28, "Departure Date")
                .SetCellValue(1, 29, "Arrival Date")
                .SetCellValue(1, 30, "Connected Document")
                .SetCellValue(1, 31, "Pax Remarks")
                .SetCellValue(1, 32, "Doc Status ID")
                .SetCellValue(1, 33, "Cancels Docs")
                .SetCellValue(1, 34, "Sertvices Description")
                .SetCellValue(1, 35, "Client Team")
                .SetCellValue(1, 36, "Sell")
                .SetCellValue(1, 37, "Buy")
                .SetCellValue(1, 38, "Profit")
                .SetCellValue(1, 39, "PaxCount+-")
                .SetCellValue(1, 40, "01-Label")
                .SetCellValue(1, 41, "01-Mandatory")
                .SetCellValue(1, 42, "01-BookedBy (G)")
                .SetCellValue(1, 43, "02-Label")
                .SetCellValue(1, 44, "02-Mandatory")
                .SetCellValue(1, 45, "02-Office")
                .SetCellValue(1, 46, "04-Label")
                .SetCellValue(1, 47, "04-Mandatory")
                .SetCellValue(1, 48, "04-ReasonForTravel")
                .SetCellValue(1, 49, "05-Label")
                .SetCellValue(1, 50, "05-Mandatory")
                .SetCellValue(1, 51, "05-CostCentre")
                .SetCellValue(1, 52, "06-Label")
                .SetCellValue(1, 53, "06-Mandatory")
                .SetCellValue(1, 54, "06-Savings (G)")
                .SetCellValue(1, 55, "07-Label")
                .SetCellValue(1, 56, "07-Mandatory")
                .SetCellValue(1, 57, "07-Losses (G)")
                .SetCellValue(1, 58, "08-Label")
                .SetCellValue(1, 59, "08-Mandatory")
                .SetCellValue(1, 60, "08-Savings/Losses Reason (G)")
                .SetCellValue(1, 61, "09-Label")
                .SetCellValue(1, 62, "09-Mandatory")
                .SetCellValue(1, 63, "09-Travel Definition (G)")
                .SetCellValue(1, 64, "11-Label")
                .SetCellValue(1, 65, "11-Mandatory")
                .SetCellValue(1, 66, "11-RequisitionNumber")
                .SetCellValue(1, 67, "12-Label")
                .SetCellValue(1, 68, "12-Mandatory")
                .SetCellValue(1, 69, "12-Passenger ID (G)")
                .SetCellValue(1, 70, "13-Label")
                .SetCellValue(1, 71, "13-Mandatory")
                .SetCellValue(1, 72, "13-OPT (G)")
                .SetCellValue(1, 73, "14-Label")
                .SetCellValue(1, 74, "14-Mandatory")
                .SetCellValue(1, 75, "14-TRID-MarineFare")
                .SetCellValue(1, 76, "15-Label")
                .SetCellValue(1, 77, "15-Mandatory")
                .SetCellValue(1, 78, "15-AccountCode")
                .SetCellValue(1, 79, "16-Label")
                .SetCellValue(1, 80, "16-Mandatory")
                .SetCellValue(1, 81, "16-Rank")
                .SetCellValue(1, 82, "17-Label")
                .SetCellValue(1, 83, "17-Mandatory")
                .SetCellValue(1, 84, "17-Nationality")
                .SetCellValue(1, 85, "18-Label")
                .SetCellValue(1, 86, "18-Mandatory")
                .SetCellValue(1, 87, "18-REF1")
                .SetCellValue(1, 88, "19-Label")
                .SetCellValue(1, 89, "19-Mandatory")
                .SetCellValue(1, 90, "19-REF2")
                .SetCellValue(1, 91, "20-Label")
                .SetCellValue(1, 92, "20-Mandatory")
                .SetCellValue(1, 93, "20-REF3")
                .SetCellValue(1, 94, "21-Label")
                .SetCellValue(1, 95, "21-Mandatory")
                .SetCellValue(1, 96, "21-REF4")
                .SetCellValue(1, 97, "22-Label")
                .SetCellValue(1, 98, "22-Mandatory")
                .SetCellValue(1, 99, "22-REF5")
                .SetCellValue(1, 100, "23-Label")
                .SetCellValue(1, 101, "23-Mandatory")
                .SetCellValue(1, 102, "23-REF6")
                .SetCellValue(1, 103, "24-Label")
                .SetCellValue(1, 104, "24-Mandatory")
                .SetCellValue(1, 105, "24-REF7")
                .SetCellValue(1, 106, "25-Label")
                .SetCellValue(1, 107, "25-Mandatory")
                .SetCellValue(1, 108, "25-REF8")
                .SetCellValue(1, 109, "26-Label")
                .SetCellValue(1, 110, "26-Mandatory")
                .SetCellValue(1, 111, "26-REF9")
                .SetCellValue(1, 112, "27-Label")
                .SetCellValue(1, 113, "27-Mandatory")
                .SetCellValue(1, 114, "27-REF10")
                .SetCellValue(1, 115, "28-Label")
                .SetCellValue(1, 116, "28-Mandatory")
                .SetCellValue(1, 117, "28-REF11")
                .SetCellValue(1, 118, "29-Label")
                .SetCellValue(1, 119, "29-Mandatory")
                .SetCellValue(1, 120, "29-REF12")
                .SetCellValue(1, 121, "30-Label")
                .SetCellValue(1, 122, "30-Mandatory")
                .SetCellValue(1, 123, "30-REF13")
                .SetCellValue(1, 124, "31-Label")
                .SetCellValue(1, 125, "31-Mandatory")
                .SetCellValue(1, 126, "31-REF14")
                .SetCellValue(1, 127, "32-Label")
                .SetCellValue(1, 128, "32-Mandatory")
                .SetCellValue(1, 129, "32-REF15")
                .SetCellValue(1, 130, "Fare Basis")
                .SetCellValue(1, 131, "Cabin Class")
                .SetCellValue(1, 132, "Duration")
                .SetCellValue(1, 133, "CO2")

                .SetCellStyle(1, 1, 1, 133, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To 10
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(13) <> 0 Then
                        For j = 11 To 14
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                    End If
                    For j = 15 To 31
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If CInt(mdsDataSet.Tables(0).Rows(i).Item(31)) = 43 Then
                        .SetCellValue(xlINVCount, 33, "Cancelled")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 129, xlStyleCancelled)
                    ElseIf CStr(mdsDataSet.Tables(0).Rows(i).Item(32)) <> "" Then
                        .SetCellValue(xlINVCount, 33, $"Cancels {CStr(mdsDataSet.Tables(0).Rows(i).Item(32))}")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 129, xlStyleCancelled)
                    End If
                    For j = 33 To 38
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    For j = 39 To 126 Step 3
                        If mdsDataSet.Tables(0).Rows(i).Item(j) = "" And mdsDataSet.Tables(0).Rows(i).Item(j + 2) = "" Then
                            .SetCellStyle(xlINVCount, j + 1, xlINVCount, j + 3, xlStyleGreyOut)
                        Else
                            If mdsDataSet.Tables(0).Rows(i).Item(j + 1) = "Mandatory" Then
                                .SetCellStyle(xlINVCount, j + 1, xlINVCount, j + 3, xlStyleMandatory)
                            End If
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            .SetCellValue(xlINVCount, j + 2, mdsDataSet.Tables(0).Rows(i).Item(j + 1))
                            .SetCellValue(xlINVCount, j + 3, mdsDataSet.Tables(0).Rows(i).Item(j + 2))
                        End If
                    Next
                    .SetCellValue(xlINVCount, 34, mdsDataSet.Tables(0).Rows(i).Item(33))
                    .SetCellValue(xlINVCount, 35, mdsDataSet.Tables(0).Rows(i).Item(34))
                    .SetCellValue(xlINVCount, 130, mdsDataSet.Tables(0).Rows(i).Item(129))
                    .SetCellValue(xlINVCount, 131, mdsDataSet.Tables(0).Rows(i).Item(130))
                    .SetCellValue(xlINVCount, 132, mdsDataSet.Tables(0).Rows(i).Item(131).ToString)
                    .SetCellValue(xlINVCount, 133, mdsDataSet.Tables(0).Rows(i).Item(132).ToString)

                    If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 130, xlStyleOmit)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 130, xlStyleVoid)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 130, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 133)
                For ic = 40 To 127 Step 3
                    .GroupColumns(ic, ic + 1)
                    .CollapseColumns(ic + 2)
                Next

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E47_DailyProfitReportTotalsOnly(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        ' Data SS Description
        '	00	  Tots
        '	01	  ClientGroupDescription
        '	02 01 ClientCode
        '	03 02 ClientName
        'CURRENT AIR
        '	04 03 NetPayableAir
        '	05 04 NetBuyAIR
        '	06 05   IW5AIR
        '	07 06   IW6AIR
        '	08 07   IW7AIR
        '	09 08   IW8AIR
        '	10 09   IW9AIR
        '                     IW11AIR
        '   11 10   IW10AIR
        '	12 11 IWAIR
        '	13 12 ProfitAir
        '	14 13 PaxAIR
        '	15 14 ProfitPerPaxAir
        'CURRENT SERVICES
        '	16 15 NetPayableServices
        '	17 16 NetBuyServices
        '	18 17   IW5Services
        '	19 18   IW6Services
        '	20 19   IW7Services
        '	21 20   IW8Services
        '	22 21   IW9Services
        '                     IW11Services
        '	23 22   IW10Services
        '	24 23 IWServices
        '	25 24 ProfitServices
        '	26 25 PaxServices
        '	27 26 ProfitPerPaxServices
        'CURRENT TOTAL
        '	28 27 NetPayable
        '	29 28 NetBuy
        '	30 29   IW5
        '	31 30   IW6
        '	32 31   IW7
        '	33 32   IW8
        '	34 33   IW9
        '                     IW11
        '	35 34   IW10
        '	36 35	IW
        '	37 36 Profit
        '	38 37 Pax
        '	39 38 ProfitPerPax
        'YTD AIR
        '	40	  NetPayableYTDAir
        '	41	  NetBuyYTDAIR
        '	42	    IW5YTDAIR
        '	43	    IW6YTDAIR
        '	44	    IW7YTDAIR
        '	45	    IW8YTDAIR
        '	46	    IW9YTDAIR
        '                     IW11YTDAIR
        '	47	    IW10YTDAIR
        '	48	  IWYTDAIR
        '	49	  ProfitYTDAir
        '	50	  PaxYTDAIR
        '	51	  ProfitPerPaxYTDAir
        'YTD SERVICES
        '	52	  NetPayableYTDServices
        '	53	  NetBuyYTDServices
        '	54	    IW5YTDServices
        '	55	    IW6YTDServices
        '	56	    IW7YTDServices
        '	57	    IW8YTDServices
        '	58	    IW9YTDServices
        '                     IW11YTDServices
        '	59	    IW10YTDServices
        '	60	  IWYTDServices
        '	61	  ProfitYTDServices
        '	62	  PaxYTDServices
        '	63	  ProfitPerPaxYTDServices
        'YTD TOTAL
        '	64	  NetPayableYTD
        '	65	  NetBuyYTD
        '	66	    IW5YTD
        '	67	    IW6YTD
        '	68	    IW7YTD
        '	69	    IW8YTD
        '	70	    IW9YTD
        '                     IW11YTD
        '	71	    IW10YTD
        '	72	  IWYTD
        '	73	  ProfitYTD
        '	74	  PaxYTD
        '	75 39 ProfitPerPaxYTD

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Totals")
                .FreezePanes(3, 2)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleIW5 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleIW5.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleIW5.Fill.SetPatternForegroundColor(Color.LightGray)


                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 6, xlNumStyle)
                .SetColumnStyle(8, xlNumStyle)
                .SetColumnStyle(10, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(9, xlNumStyle)

                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Total")
                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")
                .SetCellValue(3, 3, "Net Payable")
                .SetCellValue(3, 4, "Net Buy")
                .SetCellValue(3, 5, "IW")
                .SetCellValue(3, 6, "Profit")
                .SetCellValue(3, 7, "Total Pax")
                .SetCellValue(3, 8, "IW5")
                .SetCellValue(3, 9, "IW5 Pax")
                .SetCellValue(3, 10, "IW5/Pax")

                .SetCellStyle(1, 1, 3, 10, xlStyleHeader)
                .MergeWorksheetCells(1, 3, 1, 10)
                .MergeWorksheetCells(2, 3, 2, 10)
                If mReport.BooleanOption1 = 0 Then
                    .SetCellValue(1, 1, "INCLUDING OMIT")
                    .SetCellStyle(1, 1, xlStyleOMIT)
                End If
                xlINVCount = 3
                Dim pTotals(100) As Decimal

                For ii = 0 To mdsDataSet.Tables(0).Rows.Count - 1

                    If mdsDataSet.Tables(0).Rows(ii).Item(2) <> "" Then
                        If mdsDataSet.Tables(0).Rows(ii).Item(30) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(31) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(32) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(39) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(40) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(41) <> 0 Or
                                mdsDataSet.Tables(0).Rows(ii).Item(100) <> 0 Then
                            ' Detail rows
                            xlINVCount += 1
                            .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(ii).Item(2))
                            .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(ii).Item(3))
                            .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(ii).Item(30))
                            .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(ii).Item(31))
                            .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(ii).Item(39))
                            .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(ii).Item(40))
                            .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(ii).Item(41))
                            .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(ii).Item(32))
                            .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(ii).Item(100))
                            If mdsDataSet.Tables(0).Rows(ii).Item(100) <> 0 Then
                                .SetCellValue(xlINVCount, 10, CDec(mdsDataSet.Tables(0).Rows(ii).Item(32) / mdsDataSet.Tables(0).Rows(ii).Item(100)))
                                If mdsDataSet.Tables(0).Rows(ii).Item(39) <> mdsDataSet.Tables(0).Rows(ii).Item(32) Then
                                    .SetCellValue(xlINVCount, 11, "IW5 difference " & Format(mdsDataSet.Tables(0).Rows(ii).Item(39) - mdsDataSet.Tables(0).Rows(ii).Item(32), "#,##0.00"))
                                End If
                            End If
                            pTotals(30) += mdsDataSet.Tables(0).Rows(ii).Item(30)
                            pTotals(31) += mdsDataSet.Tables(0).Rows(ii).Item(31)
                            pTotals(32) += mdsDataSet.Tables(0).Rows(ii).Item(32)
                            pTotals(39) += mdsDataSet.Tables(0).Rows(ii).Item(39)
                            pTotals(40) += mdsDataSet.Tables(0).Rows(ii).Item(40)
                            pTotals(41) += mdsDataSet.Tables(0).Rows(ii).Item(41)
                            pTotals(100) += mdsDataSet.Tables(0).Rows(ii).Item(100)
                        End If
                    End If
                Next
                .Sort(4, 1, xlINVCount, 11, 6, False)
                .SetCellStyle(4, 8, xlINVCount, 10, xlStyleIW5)

                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                .SetCellValue(xlINVCount + 1, 3, pTotals(30))
                .SetCellValue(xlINVCount + 1, 4, pTotals(31))
                .SetCellValue(xlINVCount + 1, 5, pTotals(39))
                .SetCellValue(xlINVCount + 1, 6, pTotals(40))
                .SetCellValue(xlINVCount + 1, 7, pTotals(41))
                .SetCellValue(xlINVCount + 1, 8, pTotals(32))
                .SetCellValue(xlINVCount + 1, 9, pTotals(100))
                If pTotals(100) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 10, CDec(pTotals(32) / pTotals(100)))
                End If
                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 10, xlStyleHeader)

                .AutoFitColumn(0, 11)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E48_DailyProfitReportTotalsPerInvoice(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Totals")
                .FreezePanes(3, 2)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOmitItem As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmitItem.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmitItem.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleIW5 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleIW5.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleIW5.Fill.SetPatternForegroundColor(Color.LightGray)


                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(7, 10, xlNumStyle)
                .SetColumnStyle(12, xlNumStyle)
                .SetColumnStyle(14, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(11, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                xlNumStyle.FormatCode = "dd/MM/yyyy"
                .SetColumnStyle(5, xlNumStyle)

                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Total")
                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")

                .SetCellValue(3, 3, "Doc Type")
                .SetCellValue(3, 4, "Doc Number")
                .SetCellValue(3, 5, "Doc Date")
                .SetCellValue(3, 6, "Omit")

                .SetCellValue(3, 7, "Net Payable")
                .SetCellValue(3, 8, "Net Buy")
                .SetCellValue(3, 9, "IW")
                .SetCellValue(3, 10, "Profit")
                .SetCellValue(3, 11, "Total Pax")
                .SetCellValue(3, 12, "IW5")
                .SetCellValue(3, 13, "IW5 Pax")
                .SetCellValue(3, 14, "IW5/Pax")
                .SetCellValue(3, 15, "Notes")
                .SetCellStyle(1, 1, 3, 15, xlStyleHeader)
                .MergeWorksheetCells(1, 3, 1, 15)
                .MergeWorksheetCells(2, 3, 2, 15)
                If mReport.BooleanOption1 = 0 Then
                    .SetCellValue(1, 1, "INCLUDING OMIT")
                    .SetCellStyle(1, 1, xlStyleOMIT)
                End If
                xlINVCount = 3
                Dim pTotals(100) As Decimal

                For ii = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    ' Detail rows
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(ii).Item(2))
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(ii).Item(3))
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(ii).Item(4))
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(ii).Item(5))
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(ii).Item(6))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(ii).Item(7))


                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(ii).Item(34))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(ii).Item(35))
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(ii).Item(43))
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(ii).Item(44))
                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(ii).Item(45))
                    If mdsDataSet.Tables(0).Rows(ii).Item(36) <> 0 Then
                        .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(ii).Item(36))
                    End If
                    If mdsDataSet.Tables(0).Rows(ii).Item(52) <> 0 Then
                        .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(ii).Item(52))
                    End If
                    If mdsDataSet.Tables(0).Rows(ii).Item(52) <> 0 Then
                        .SetCellValue(xlINVCount, 14, CDec(mdsDataSet.Tables(0).Rows(ii).Item(36) / mdsDataSet.Tables(0).Rows(ii).Item(52)))
                    End If
                    If mdsDataSet.Tables(0).Rows(ii).Item(43) <> mdsDataSet.Tables(0).Rows(ii).Item(36) Then
                        .SetCellValue(xlINVCount, 15, "IW5 difference " & Format(mdsDataSet.Tables(0).Rows(ii).Item(43) - mdsDataSet.Tables(0).Rows(ii).Item(36), "#,##0.00"))
                    End If
                    pTotals(34) += mdsDataSet.Tables(0).Rows(ii).Item(34)
                    pTotals(35) += mdsDataSet.Tables(0).Rows(ii).Item(35)
                    pTotals(43) += mdsDataSet.Tables(0).Rows(ii).Item(43)
                    pTotals(44) += mdsDataSet.Tables(0).Rows(ii).Item(44)
                    pTotals(45) += mdsDataSet.Tables(0).Rows(ii).Item(45)
                    pTotals(36) += mdsDataSet.Tables(0).Rows(ii).Item(36)
                    pTotals(52) += mdsDataSet.Tables(0).Rows(ii).Item(52)
                    If mdsDataSet.Tables(0).Rows(ii).Item(7) = 1 Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 14, xlStyleOmitItem)
                    End If
                Next
                .Sort(4, 1, xlINVCount, 15, 4, True)
                .Sort(4, 1, xlINVCount, 15, 3, True)
                .Sort(4, 1, xlINVCount, 15, 1, True)
                .SetCellStyle(4, 12, xlINVCount, 15, xlStyleIW5)

                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                .SetCellValue(xlINVCount + 1, 7, pTotals(34))
                .SetCellValue(xlINVCount + 1, 8, pTotals(35))
                .SetCellValue(xlINVCount + 1, 9, pTotals(43))
                .SetCellValue(xlINVCount + 1, 10, pTotals(44))
                .SetCellValue(xlINVCount + 1, 11, pTotals(45))
                .SetCellValue(xlINVCount + 1, 12, pTotals(36))
                .SetCellValue(xlINVCount + 1, 13, pTotals(52))
                If pTotals(52) <> 0 Then
                    .SetCellValue(xlINVCount + 1, 14, CDec(pTotals(36) / pTotals(52)))
                End If
                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 14, xlStyleHeader)

                .AutoFitColumn(0, 15)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E15_DailyProfitReport_OLD_2019_01_18(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        ' Data SS Description
        '	00	  Tots
        '	01	  ClientGroupDescription
        '	02 01 ClientCode
        '	03 02 ClientName
        'CURRENT AIR
        '	04 03 NetPayableAir
        '	05 04 NetBuyAIR
        '	06 05   IW5AIR
        '	07 06   IW6AIR
        '	08 07   IW7AIR
        '	09 08   IW8AIR
        '	10 09   IW9AIR
        '	11 10 IWAIR
        '	12 11 ProfitAir
        '	13 12 PaxAIR
        '	14 13 ProfitPerPaxAir
        'CURRENT SERVICES
        '	15 14 NetPayableServices
        '	16 15 NetBuyServices
        '	17 16   IW5Services
        '	18 17   IW6Services
        '	19 18   IW7Services
        '	20 19   IW8Services
        '	21 20   IW9Services
        '	22 21 IWServices
        '	23 22 ProfitServices
        '	24 23 PaxServices
        '	25 24 ProfitPerPaxServices
        'CURRENT TOTAL
        '	26 25 NetPayable
        '	27 26 NetBuy
        '	28 27   IW5
        '	29 28   IW6
        '	30 29   IW7
        '	31 30   IW8
        '	32 31   IW9
        '	33 32	IW
        '	34 33 Profit
        '	35 34 Pax
        '	36 35 ProfitPerPax
        'YTD AIR
        '	37	  NetPayableYTDAir
        '	38	  NetBuyYTDAIR
        '	39	    IW5YTDAIR
        '	40	    IW6YTDAIR
        '	41	    IW7YTDAIR
        '	42	    IW8YTDAIR
        '	43	    IW9YTDAIR
        '	44	  IWYTDAIR
        '	45	  ProfitYTDAir
        '	46	  PaxYTDAIR
        '	47	  ProfitPerPaxYTDAir
        'YTD SERVICES
        '	48	  NetPayableYTDServices
        '	49	  NetBuyYTDServices
        '	50	    IW5YTDServices
        '	51	    IW6YTDServices
        '	52	    IW7YTDServices
        '	53	    IW8YTDServices
        '	54	    IW9YTDServices
        '	55	  IWYTDServices
        '	56	  ProfitYTDServices
        '	57	  PaxYTDServices
        '	58	  ProfitPerPaxYTDServices
        'YTD TOTAL
        '	59	  NetPayableYTD
        '	60	  NetBuyYTD
        '	61	    IW5YTD
        '	62	    IW6YTD
        '	63	    IW7YTD
        '	64	    IW8YTD
        '	65	    IW9YTD
        '	66	  IWYTD
        '	67	  ProfitYTD
        '	68	  PaxYTD
        '	69 36 ProfitPerPaxYTD

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 2)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 36, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(12, xlNumStyle)
                .SetColumnStyle(23, xlNumStyle)
                .SetColumnStyle(34, xlNumStyle)
                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 14, "Other Services")
                .SetCellValue(2, 25, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")

                For i = 3 To 25 Step 11
                    .SetCellValue(3, i, "Net Payable")
                    .SetCellValue(3, i + 1, "Net Buy")
                    .SetCellValue(3, i + 2, "IW5")
                    .SetCellValue(3, i + 3, "IW6")
                    .SetCellValue(3, i + 4, "IW7")
                    .SetCellValue(3, i + 5, "IW8")
                    .SetCellValue(3, i + 6, "IW9")
                    .SetCellValue(3, i + 7, "IW")
                    .SetCellValue(3, i + 8, "Profit")
                    .SetCellValue(3, i + 9, "Pax")
                    .SetCellValue(3, i + 10, "Profit Per Pax")
                Next

                .SetCellValue(2, 36, "YTD")
                .SetCellValue(3, 36, "Profit Per Pax")

                .SetCellStyle(1, 3, 3, 36, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 35)
                .MergeWorksheetCells(2, 3, 2, 13)
                .MergeWorksheetCells(2, 14, 2, 24)
                .MergeWorksheetCells(2, 25, 2, 35)


                xlINVCount = 3
                Dim pTotals(69) As Decimal
                Dim pOtherClients(69) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "1" And (pTopClients = 0 Or pTopClientsCount < pTopClients) Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = "2" And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 36
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    .SetCellValue(xlINVCount, 36, mdsDataSet.Tables(0).Rows(ii).Item(69))
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 36, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(36) < mdsDataSet.Tables(0).Rows(ii).Item(69) Then
                                        .SetCellStyle(xlINVCount, 35, xlStyleDetailNeg)
                                    End If
                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        xlINVCount += 1
                        For j = 2 To 36
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pTotals(67) += mdsDataSet.Tables(0).Rows(i).Item(67)
                        pTotals(68) += mdsDataSet.Tables(0).Rows(i).Item(68)
                        .SetCellValue(xlINVCount, 36, mdsDataSet.Tables(0).Rows(i).Item(69))
                        If mdsDataSet.Tables(0).Rows(i).Item(36) < mdsDataSet.Tables(0).Rows(i).Item(69) Then
                            .SetCellStyle(xlINVCount, 35, xlStyleNegative)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 36
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pOtherClients(67) += mdsDataSet.Tables(0).Rows(i).Item(67)
                        pOtherClients(68) += mdsDataSet.Tables(0).Rows(i).Item(68)
                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(13) <> 0 Then
                        pOtherClients(14) = pOtherClients(12) / pOtherClients(13)
                    Else
                        pOtherClients(14) = 0
                    End If
                    If pOtherClients(24) <> 0 Then
                        pOtherClients(25) = pOtherClients(23) / pOtherClients(24)
                    Else
                        pOtherClients(25) = 0
                    End If
                    If pOtherClients(35) <> 0 Then
                        pOtherClients(36) = pOtherClients(34) / pOtherClients(35)
                    Else
                        pOtherClients(36) = 0
                    End If
                    If pOtherClients(68) <> 0 Then
                        pOtherClients(69) = pOtherClients(67) / pOtherClients(68)
                    Else
                        pOtherClients(69) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 36
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    pTotals(67) += pOtherClients(67)
                    pTotals(68) += pOtherClients(68)
                    .SetCellValue(xlINVCount, 36, pOtherClients(69))
                    If pOtherClients(36) < pOtherClients(69) Then
                        .SetCellStyle(xlINVCount, 35, xlStyleNegative)
                    End If

                End If
                If pTotals(13) <> 0 Then
                    pTotals(14) = pTotals(12) / pTotals(13)
                Else
                    pTotals(14) = 0
                End If
                If pTotals(24) <> 0 Then
                    pTotals(25) = pTotals(23) / pTotals(24)
                Else
                    pTotals(25) = 0
                End If
                If pTotals(35) <> 0 Then
                    pTotals(36) = pTotals(34) / pTotals(35)
                Else
                    pTotals(36) = 0
                End If
                If pTotals(68) <> 0 Then
                    pTotals(69) = pTotals(67) / pTotals(68)
                Else
                    pTotals(69) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 36
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 36, pTotals(69))
                If pTotals(36) < pTotals(69) Then
                    .SetCellStyle(xlINVCount + 1, 35, xlStyleNegative)
                End If

                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 36, xlStyleHeader)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 13, xlINVCount + 1, 13, xlStyleHeader)
                .SetCellStyle(3, 24, xlINVCount + 1, 24, xlStyleHeader)
                .SetCellStyle(3, 35, xlINVCount + 1, 35, xlStyleHeader)
                .SetCellStyle(3, 36, xlINVCount + 1, 36, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 36, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(12, xlNumStyle)
                .SetColumnStyle(23, xlNumStyle)
                .SetColumnStyle(34, xlNumStyle)

                .SetColumnStyle(5, 9, xlStyleDetail)
                .SetColumnStyle(16, 20, xlStyleDetail)
                .SetColumnStyle(27, 31, xlStyleDetail)

                .AutoFitColumn(0, 36)
                .GroupColumns(5, 9)
                .CollapseColumns(10)
                .GroupColumns(16, 20)
                .CollapseColumns(21)
                .GroupColumns(27, 31)
                .CollapseColumns(32)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E15_DailyProfitReportWithoutIWAnalysis(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 0)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 21, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                .SetColumnStyle(19, xlNumStyle)
                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 9, "Other Services")
                .SetCellValue(2, 15, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")
                .SetCellValue(3, 3, "Net Payable")
                .SetCellValue(3, 4, "Net Buy")
                .SetCellValue(3, 5, "IW")
                .SetCellValue(3, 6, "Profit")
                .SetCellValue(3, 7, "Pax")
                .SetCellValue(3, 8, "Profit Per Pax")
                .SetCellValue(3, 9, "Net Payable")
                .SetCellValue(3, 10, "Net Buy")
                .SetCellValue(3, 11, "IW")
                .SetCellValue(3, 12, "Profit")
                .SetCellValue(3, 13, "Pax")
                .SetCellValue(3, 14, "Profit Per Pax")
                .SetCellValue(3, 15, "Net Payable")
                .SetCellValue(3, 16, "Net Buy")
                .SetCellValue(3, 17, "IW")
                .SetCellValue(3, 18, "Profit")
                .SetCellValue(3, 19, "Pax")
                .SetCellValue(3, 20, "Profit Per Pax")

                .SetCellValue(2, 21, "YTD")
                .SetCellValue(3, 21, "Profit Per Pax")

                .SetCellStyle(1, 3, 3, 21, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 20)
                .MergeWorksheetCells(2, 3, 2, 8)
                .MergeWorksheetCells(2, 9, 2, 14)
                .MergeWorksheetCells(2, 15, 2, 20)


                xlINVCount = 3
                Dim pTotals(39) As Decimal
                Dim pOtherClients(39) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = 1 And (pTopClients = 0 Or pTopClientsCount < pTopClients) Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = 2 And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 21
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    .SetCellValue(xlINVCount, 21, mdsDataSet.Tables(0).Rows(ii).Item(39))
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 21, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(21) < mdsDataSet.Tables(0).Rows(ii).Item(39) Then
                                        .SetCellStyle(xlINVCount, 20, xlStyleDetailNeg)
                                    End If
                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        xlINVCount += 1
                        For j = 2 To 21
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pTotals(37) += mdsDataSet.Tables(0).Rows(i).Item(37)
                        pTotals(38) += mdsDataSet.Tables(0).Rows(i).Item(38)
                        .SetCellValue(xlINVCount, 21, mdsDataSet.Tables(0).Rows(i).Item(39))
                        If mdsDataSet.Tables(0).Rows(i).Item(21) < mdsDataSet.Tables(0).Rows(i).Item(39) Then
                            .SetCellStyle(xlINVCount, 20, xlStyleNegative)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 21
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        pOtherClients(37) += mdsDataSet.Tables(0).Rows(i).Item(37)
                        pOtherClients(38) += mdsDataSet.Tables(0).Rows(i).Item(38)
                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(8) <> 0 Then
                        pOtherClients(9) = pOtherClients(7) / pOtherClients(8)
                    Else
                        pOtherClients(9) = 0
                    End If
                    If pOtherClients(14) <> 0 Then
                        pOtherClients(15) = pOtherClients(13) / pOtherClients(14)
                    Else
                        pOtherClients(15) = 0
                    End If
                    If pOtherClients(20) <> 0 Then
                        pOtherClients(21) = pOtherClients(19) / pOtherClients(20)
                    Else
                        pOtherClients(21) = 0
                    End If
                    If pOtherClients(38) <> 0 Then
                        pOtherClients(39) = pOtherClients(37) / pOtherClients(38)
                    Else
                        pOtherClients(39) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 21
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    pTotals(37) += pOtherClients(37)
                    pTotals(38) += pOtherClients(38)
                    .SetCellValue(xlINVCount, 21, pOtherClients(39))
                    If pOtherClients(21) < pOtherClients(39) Then
                        .SetCellStyle(xlINVCount, 20, xlStyleNegative)
                    End If

                End If
                If pTotals(8) <> 0 Then
                    pTotals(9) = pTotals(7) / pTotals(8)
                Else
                    pTotals(9) = 0
                End If
                If pTotals(14) <> 0 Then
                    pTotals(15) = pTotals(13) / pTotals(14)
                Else
                    pTotals(15) = 0
                End If
                If pTotals(20) <> 0 Then
                    pTotals(21) = pTotals(19) / pTotals(20)
                Else
                    pTotals(21) = 0
                End If
                If pTotals(38) <> 0 Then
                    pTotals(39) = pTotals(37) / pTotals(38)
                Else
                    pTotals(39) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 21
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                .SetCellValue(xlINVCount + 1, 21, pTotals(39))
                If pTotals(21) < pTotals(39) Then
                    .SetCellStyle(xlINVCount + 1, 20, xlStyleNegative)
                End If

                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 21, xlStyleHeader)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 8, xlINVCount + 1, 8, xlStyleHeader)
                .SetCellStyle(3, 14, xlINVCount + 1, 14, xlStyleHeader)
                .SetCellStyle(3, 20, xlINVCount + 1, 20, xlStyleHeader)
                .SetCellStyle(3, 21, xlINVCount + 1, 21, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 21, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                .SetColumnStyle(19, xlNumStyle)

                .AutoFitColumn(0, 21)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E16_DailyProfitReportWithYTD(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try



            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(3, 0)

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleDetailNeg As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetailNeg.Font.FontColor = Color.Red
                xlStyleDetailNeg.Font.Italic = True

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 38, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                .SetColumnStyle(19, xlNumStyle)

                .SetColumnStyle(25, xlNumStyle)
                .SetColumnStyle(31, xlNumStyle)
                .SetColumnStyle(37, xlNumStyle)

                .SetCellValue(1, 3, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(1, 21, Format(DateSerial(mReport.Date1From.Year, 1, 1), "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))

                .SetCellValue(2, 3, "Air Tickets")
                .SetCellValue(2, 9, "Other Services")
                .SetCellValue(2, 15, "Total")

                .SetCellValue(2, 21, "Air Tickets")
                .SetCellValue(2, 27, "Other Services")
                .SetCellValue(2, 33, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")
                ' Current month
                .SetCellValue(3, 3, "Net Payable")
                .SetCellValue(3, 4, "Net Buy")
                .SetCellValue(3, 5, "IW")
                .SetCellValue(3, 6, "Profit")
                .SetCellValue(3, 7, "Pax")
                .SetCellValue(3, 8, "Profit Per Pax")

                .SetCellValue(3, 9, "Net Payable")
                .SetCellValue(3, 10, "Net Buy")
                .SetCellValue(3, 11, "IW")
                .SetCellValue(3, 12, "Profit")
                .SetCellValue(3, 13, "Pax")
                .SetCellValue(3, 14, "Profit Per Pax")

                .SetCellValue(3, 15, "Net Payable")
                .SetCellValue(3, 16, "Net Buy")
                .SetCellValue(3, 17, "IW")
                .SetCellValue(3, 18, "Profit")
                .SetCellValue(3, 19, "Pax")
                .SetCellValue(3, 20, "Profit Per Pax")
                ' YTD
                .SetCellValue(3, 21, "Net Payable")
                .SetCellValue(3, 22, "Net Buy")
                .SetCellValue(3, 23, "IW")
                .SetCellValue(3, 24, "Profit")
                .SetCellValue(3, 25, "Pax")
                .SetCellValue(3, 26, "Profit Per Pax")

                .SetCellValue(3, 27, "Net Payable")
                .SetCellValue(3, 28, "Net Buy")
                .SetCellValue(3, 29, "IW")
                .SetCellValue(3, 30, "Profit")
                .SetCellValue(3, 31, "Pax")
                .SetCellValue(3, 32, "Profit Per Pax")

                .SetCellValue(3, 33, "Net Payable")
                .SetCellValue(3, 34, "Net Buy")
                .SetCellValue(3, 35, "IW")
                .SetCellValue(3, 36, "Profit")
                .SetCellValue(3, 37, "Pax")
                .SetCellValue(3, 38, "Profit Per Pax")

                .SetCellStyle(1, 3, 3, 38, xlStyleHeader)

                .MergeWorksheetCells(1, 3, 1, 20)
                .MergeWorksheetCells(1, 21, 1, 38)

                .MergeWorksheetCells(2, 3, 2, 8)
                .MergeWorksheetCells(2, 9, 2, 14)
                .MergeWorksheetCells(2, 15, 2, 20)

                .MergeWorksheetCells(2, 21, 2, 26)
                .MergeWorksheetCells(2, 27, 2, 32)
                .MergeWorksheetCells(2, 33, 2, 38)

                xlINVCount = 3
                Dim pTotals(39) As Decimal
                Dim pOtherClients(39) As Decimal
                Dim pTopClients As Integer = 0
                Dim pTopClientsCount As Integer = 0
                If IsNumeric(mReport.TextEntry) Then
                    pTopClients = CInt(mReport.TextEntry)
                End If
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = 1 And pTopClients = 0 Or pTopClientsCount < pTopClients Then
                        If mReport.BooleanOption1 And mdsDataSet.Tables(0).Rows(i).Item(1) <> "" Then
                            Dim pFirstGroupRow As Integer = xlINVCount + 1
                            Dim pLastGroupRow As Integer = xlINVCount + 1
                            For ii As Integer = i To mdsDataSet.Tables(0).Rows.Count - 1
                                If mdsDataSet.Tables(0).Rows(ii).Item(0) = 2 And mdsDataSet.Tables(0).Rows(ii).Item(1) = mdsDataSet.Tables(0).Rows(i).Item(1) Then
                                    xlINVCount += 1
                                    For jj As Integer = 2 To 39
                                        .SetCellValue(xlINVCount, jj - 1, mdsDataSet.Tables(0).Rows(ii).Item(jj))
                                    Next
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleDetail)
                                    If mdsDataSet.Tables(0).Rows(ii).Item(21) < mdsDataSet.Tables(0).Rows(ii).Item(39) Then
                                        .SetCellStyle(xlINVCount, 20, xlStyleDetailNeg)
                                    End If
                                    pLastGroupRow = xlINVCount
                                End If
                            Next
                            .GroupRows(pFirstGroupRow, pLastGroupRow)
                            .CollapseRows(pLastGroupRow + 1)
                        End If
                        xlINVCount += 1
                        For j = 2 To 39
                            .SetCellValue(xlINVCount, j - 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            If j > 3 Then
                                pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                        If mdsDataSet.Tables(0).Rows(i).Item(21) < mdsDataSet.Tables(0).Rows(i).Item(39) Then
                            .SetCellStyle(xlINVCount, 20, xlStyleNegative)
                        End If
                        pTopClientsCount += 1
                    Else
                        For j = 2 To 39
                            If j > 3 Then
                                pOtherClients(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                            End If
                        Next
                    End If
                Next
                If pTopClients > 0 And pTopClientsCount >= pTopClients Then
                    If pOtherClients(8) <> 0 Then
                        pOtherClients(9) = pOtherClients(7) / pOtherClients(8)
                    Else
                        pOtherClients(9) = 0
                    End If
                    If pOtherClients(14) <> 0 Then
                        pOtherClients(15) = pOtherClients(13) / pOtherClients(14)
                    Else
                        pOtherClients(15) = 0
                    End If
                    If pOtherClients(20) <> 0 Then
                        pOtherClients(21) = pOtherClients(19) / pOtherClients(20)
                    Else
                        pOtherClients(21) = 0
                    End If
                    If pOtherClients(38) <> 0 Then
                        pOtherClients(39) = pOtherClients(37) / pOtherClients(38)
                    Else
                        pOtherClients(39) = 0
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, "Other Clients")
                    For j = 4 To 39
                        .SetCellValue(xlINVCount, j - 1, pOtherClients(j))
                        pTotals(j) += pOtherClients(j)
                    Next
                    If pOtherClients(21) < pOtherClients(39) Then
                        .SetCellStyle(xlINVCount, 20, xlStyleNegative)
                    End If

                End If
                If pTotals(8) <> 0 Then
                    pTotals(9) = pTotals(7) / pTotals(8)
                Else
                    pTotals(9) = 0
                End If
                If pTotals(14) <> 0 Then
                    pTotals(15) = pTotals(13) / pTotals(14)
                Else
                    pTotals(15) = 0
                End If
                If pTotals(20) <> 0 Then
                    pTotals(21) = pTotals(19) / pTotals(20)
                Else
                    pTotals(21) = 0
                End If
                If pTotals(38) <> 0 Then
                    pTotals(39) = pTotals(37) / pTotals(38)
                Else
                    pTotals(39) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 4 To 39
                    .SetCellValue(xlINVCount + 1, j - 1, pTotals(j))
                Next
                If pTotals(21) < pTotals(39) Then
                    .SetCellStyle(xlINVCount + 1, 20, xlStyleNegative)
                End If

                .SetCellStyle(xlINVCount + 1, 2, xlINVCount + 1, 38, xlStyleHeader)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, 8, xlINVCount + 1, 8, xlStyleHeader)
                .SetCellStyle(3, 14, xlINVCount + 1, 14, xlStyleHeader)
                .SetCellStyle(3, 20, xlINVCount + 1, 20, xlStyleHeader)
                .SetCellStyle(3, 26, xlINVCount + 1, 26, xlStyleHeader)
                .SetCellStyle(3, 32, xlINVCount + 1, 32, xlStyleHeader)
                .SetCellStyle(3, 38, xlINVCount + 1, 38, xlStyleHeader)

                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(3, 38, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                .SetColumnStyle(19, xlNumStyle)

                .SetColumnStyle(25, xlNumStyle)
                .SetColumnStyle(31, xlNumStyle)
                .SetColumnStyle(37, xlNumStyle)

                .AutoFitColumn(0, 38)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E17_ServiceFeeAnalysis(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Report Invoices")
                .FreezePanes(1, 0)

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                Dim xlTextStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlTextStyle.FormatCode = "@"
                Dim xlDateStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlDateStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, 5, xlTextStyle)
                .SetColumnStyle(6, xlDateStyle)
                .SetColumnStyle(7, 7, xlTextStyle)
                .SetColumnStyle(8, 10, xlNumStyle)
                For i = 11 To 27 Step 2
                    .SetColumnStyle(i, xlTextStyle)
                    .SetColumnStyle(i + 1, xlNumStyle)
                Next

                .SetCellValue(1, 1, "Code")
                .SetCellValue(1, 2, "Series")
                .SetCellValue(1, 3, "Number")
                .SetCellValue(1, 4, "Code")
                .SetCellValue(1, 5, "Name")
                .SetCellValue(1, 6, "IssueDate")
                .SetCellValue(1, 7, "CurrencyCode")
                .SetCellValue(1, 8, "InvoiceServiceFee")
                .SetCellValue(1, 9, "CardServFeeAmount")
                .SetCellValue(1, 10, "SFDifference")
                .SetCellValue(1, 11, "TF Desc")
                .SetCellValue(1, 12, "TF")
                .SetCellValue(1, 13, "TF Dom Desc")
                .SetCellValue(1, 14, "TF Dom")
                .SetCellValue(1, 15, "IW5 Desc")
                .SetCellValue(1, 16, "IW5")
                .SetCellValue(1, 17, "IW6 Desc")
                .SetCellValue(1, 18, "IW6")
                .SetCellValue(1, 19, "IW7 Desc")
                .SetCellValue(1, 20, "IW7")
                .SetCellValue(1, 21, "IW8 Desc")
                .SetCellValue(1, 22, "IW8")
                .SetCellValue(1, 23, "IW9 Desc")
                .SetCellValue(1, 24, "IW9")
                .SetCellValue(1, 25, "IW10 Desc")
                .SetCellValue(1, 26, "IW10")
                .SetCellValue(1, 27, "IW11 Desc")
                .SetCellValue(1, 28, "IW11")

                .SetCellStyle(1, 1, 1, 28, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    For j = 0 To 27
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                Next
                .SetCellStyle(2, 11, xlINVCount + 1, 12, xlStyleOmit)
                .SetCellStyle(2, 15, xlINVCount + 1, 16, xlStyleOmit)
                .SetCellStyle(2, 19, xlINVCount + 1, 20, xlStyleOmit)
                .SetCellStyle(2, 23, xlINVCount + 1, 24, xlStyleOmit)
                .SetCellStyle(2, 27, xlINVCount + 1, 28, xlStyleOmit)
                .AutoFitColumn(1, 28)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E18_AirTicketSales(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 42, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(25, xlNumStyle)
                .SetColumnStyle(43, 45, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, xlNumStyle)
                .SetColumnStyle(15, xlNumStyle)
                .SetColumnStyle(36, 37, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(9, xlNumStyle)
                .SetColumnStyle(46, xlNumStyle)
                xlNumStyle.FormatCode = "0;-0;"
                .SetColumnStyle(29, xlNumStyle)

                .SetCellValue(1, 1, "Issue Date")
                .SetCellValue(1, 2, "Client Code")
                .SetCellValue(1, 3, "Client Name")
                ' if Ignore OMIT/VOID is selected, change the title so as not to show that these are valid fields
                If mReport.BooleanOption1 Then
                    .SetCellValue(1, 4, "Sell")
                    .SetCellValue(1, 5, "X")
                Else
                    .SetCellValue(1, 4, "Omit")
                    .SetCellValue(1, 5, "Void")
                End If
                .SetCellValue(1, 6, "PNR")
                .SetCellValue(1, 7, "Ticket Number")
                .SetCellValue(1, 8, "Passenger")
                .SetCellValue(1, 9, "Pax Count")
                .SetCellValue(1, 10, "Product Type")
                .SetCellValue(1, 11, "Action Type")
                .SetCellValue(1, 12, "Inv Code")
                .SetCellValue(1, 13, "Inv Series")
                .SetCellValue(1, 14, "Inv Number")
                .SetCellValue(1, 15, "Invoice Date")
                .SetCellValue(1, 16, "Vessel")
                .SetCellValue(1, 17, "Booked By")
                .SetCellValue(1, 18, "Office/Dept")
                .SetCellValue(1, 19, "Reason For Travel")
                .SetCellValue(1, 20, "Cost Centre")

                .SetCellValue(1, 21, "Requisition Number")
                .SetCellValue(1, 22, "OPT")
                .SetCellValue(1, 23, "TRID/MarineFare")
                .SetCellValue(1, 24, "Account Code")

                .SetCellValue(1, 25, "Net Payable")
                .SetCellValue(1, 26, "Verified")
                .SetCellValue(1, 27, "Remarks")
                .SetCellValue(1, 28, "Transaction Type")
                .SetCellValue(1, 29, "RegNr")
                .SetCellValue(1, 30, "Ticketing Airline")
                .SetCellValue(1, 31, "Routing")
                .SetCellValue(1, 32, "SalesPerson")
                .SetCellValue(1, 33, "Issuing Agent")
                .SetCellValue(1, 34, "Creator Agent")
                .SetCellValue(1, 35, "Reference")
                .SetCellValue(1, 36, "Departure Date")
                .SetCellValue(1, 37, "Arrival Date")
                .SetCellValue(1, 38, "Connected Document")
                .SetCellValue(1, 39, "Pax Remarks")
                .SetCellValue(1, 40, "Invoice Status")
                .SetCellValue(1, 41, "Other Services")
                .SetCellValue(1, 42, "Client Team")
                .SetCellValue(1, 43, "Sell")
                .SetCellValue(1, 44, "Buy")
                .SetCellValue(1, 45, "Profit")
                .SetCellValue(1, 46, "PaxCount+-")

                .SetCellStyle(1, 1, 1, 46, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    If Not mReport.BooleanOption1 Or (mdsDataSet.Tables(0).Rows(i).Item(3) = "" And mdsDataSet.Tables(0).Rows(i).Item(4) = "" And mdsDataSet.Tables(0).Rows(i).Item(37).ToString.Trim = "") Then

                        xlINVCount += 1
                        For j = 0 To 10
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                        ' if Ignore OMIT/VOID is selected, change contents of these columns
                        ' it's actually too much of a hassle to remove these columns and move everything two columns to the left
                        If mReport.BooleanOption1 Then
                            .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(25))
                        End If
                        If mdsDataSet.Tables(0).Rows(i).Item(13) <> 0 Then
                            For j = 11 To 14
                                .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            Next
                        End If
                        For j = 15 To 38
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                        If CInt(mdsDataSet.Tables(0).Rows(i).Item(39)) = 43 Then
                            .SetCellValue(xlINVCount, 40, "Cancelled")
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleCancelled)
                        ElseIf CStr(mdsDataSet.Tables(0).Rows(i).Item(40)) <> "" Then
                            .SetCellValue(xlINVCount, 40, $"Cancels {CStr(mdsDataSet.Tables(0).Rows(i).Item(40))}")
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleCancelled)
                        End If
                        .SetCellValue(xlINVCount, 41, mdsDataSet.Tables(0).Rows(i).Item(41))
                        .SetCellValue(xlINVCount, 42, mdsDataSet.Tables(0).Rows(i).Item(42))
                        If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleOmit)
                        End If
                        If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleVoid)
                        End If
                        If mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund" Then
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 46, xlStyleRefund)
                        End If
                        For j = 43 To 46
                            .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                    End If

                Next

                .AutoFitColumn(1, 46)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E19_DailyProfitReportInvoicesWithTicketNumber(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        ' Field numbers from SQL (0-based) (Spreadsheet column = DBITEM +1)
        Dim DBITEMAirProfitPerPax As Integer = 26
        Dim DBITEMSerProfitPerPax As Integer = 45
        Dim DBITEMTotProfitPerPax As Integer = 64

        ' 00 Client Code	
        ' 01 Client Name	
        ' 02 Invoice Date	
        ' 03 Invoice Type	
        ' 04 Invoice Number	
        ' 05 Airline	
        ' 06 Ticket Number

        ' AIR
        ' 07 Face value	
        ' 08 Taxes	
        ' 09 Commission	
        ' 10 Discount	
        ' 11 Cancellation Fee	
        ' 12 TF	
        ' 13 Net Payable	
        ' 14 Net Buy	
        ' 15 IW5	
        ' 16 IW6	
        ' 17 IW7	
        ' 18 IW8	
        ' 19 IW9	
        ' 20 IW11	
        ' 21 IW10	
        ' 22 IW	
        ' 23 Profit	
        ' 24 Pax	
        ' 25 Profit Per Pax	

        ' SERVICES
        ' 26 Face value	
        ' 27 Taxes	
        ' 28 Commission	
        ' 29 Discount	
        ' 30 Cancellation Fee	
        ' 31 TF	
        ' 32 Net Payable	
        ' 33 Net Buy	
        ' 34 IW5	
        ' 35 IW6	
        ' 36 IW7	
        ' 37 IW8	
        ' 38 IW9	
        ' 39 IW11	
        ' 40 IW10	
        ' 41 IW	
        ' 42 Profit	
        ' 43 Pax	
        ' 44 Profit Per Pax	

        ' TOTAL
        ' 45 Face value	
        ' 46 Taxes	
        ' 47 Commission	
        ' 48 Discount	
        ' 49 Cancellation Fee	
        ' 50 TF	
        ' 51 Net Payable	
        ' 52 Net Buy	
        ' 53 IW5	
        ' 54 IW6	
        ' 55 IW7	
        ' 56 IW8	
        ' 57 IW9	
        ' 58 IW11	
        ' 59 IW10	
        ' 60 IW	
        ' 61 Profit	
        ' 62 Pax	
        ' 63 Profit Per Pax	

        '--------
        ' 64 Omit

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try


            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Report Invoices")
                .FreezePanes(3, 8)

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(9, 65, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(4, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(26, xlNumStyle)
                .SetColumnStyle(45, xlNumStyle)
                .SetColumnStyle(64, xlNumStyle)
                .SetCellValue(1, 9, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(2, 9, "Air Tickets")
                .SetCellValue(2, 28, "Other Services")
                .SetCellValue(2, 47, "Total")

                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")
                .SetCellValue(3, 3, "Vessel")
                .SetCellValue(3, 4, "Invoice Date")
                .SetCellValue(3, 5, "Invoice Type")
                .SetCellValue(3, 6, "Invoice Number")
                .SetCellValue(3, 7, "Airline")
                .SetCellValue(3, 8, "Ticket Number")

                For i = 9 To 48 Step 19
                    .SetCellValue(3, i, "Face value")
                    .SetCellValue(3, i + 1, "Taxes")
                    .SetCellValue(3, i + 2, "Commission")
                    .SetCellValue(3, i + 3, "Discount")
                    .SetCellValue(3, i + 4, "Cancellation Fee")
                    .SetCellValue(3, i + 5, "TF")
                    .SetCellValue(3, i + 6, "Net Payable")
                    .SetCellValue(3, i + 7, "Net Buy")
                    .SetCellValue(3, i + 8, "IW5")
                    .SetCellValue(3, i + 9, "IW6")
                    .SetCellValue(3, i + 10, "IW7")
                    .SetCellValue(3, i + 11, "IW8")
                    .SetCellValue(3, i + 12, "IW9")
                    .SetCellValue(3, i + 13, "IW11")
                    .SetCellValue(3, i + 14, "IW10")
                    .SetCellValue(3, i + 15, "IW")
                    .SetCellValue(3, i + 16, "Profit")
                    .SetCellValue(3, i + 17, "Pax")
                    .SetCellValue(3, i + 18, "Profit Per Pax")
                Next
                .SetCellValue(3, 66, "Omit")
                .SetCellValue(3, 67, "Client Team")
                .SetCellValue(3, 68, "Customer Group")
                .SetCellStyle(1, 9, 3, 66, xlStyleHeader)

                .MergeWorksheetCells(1, 9, 1, 65)
                .MergeWorksheetCells(2, 9, 2, 27)
                .MergeWorksheetCells(2, 28, 2, 46)
                .MergeWorksheetCells(2, 47, 2, 65)

                xlINVCount = 3
                Dim pTotals(64) As Decimal
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To 64
                        If j > 7 Then
                            If mdsDataSet.Tables(0).Rows(i).Item(j) <> 0 Then
                                .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                            End If
                            pTotals(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        Else
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        End If
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(65) Then
                        .SetCellValue(xlINVCount, 66, "Omit")
                        .SetCellStyle(xlINVCount, 5, xlINVCount, 6, xlStyleOmit)
                    End If
                    .SetCellValue(xlINVCount, 67, mdsDataSet.Tables(0).Rows(i).Item(67))
                    .SetCellValue(xlINVCount, 68, mdsDataSet.Tables(0).Rows(i).Item(68))

                Next
                If pTotals(DBITEMAirProfitPerPax - 1) <> 0 Then
                    pTotals(DBITEMAirProfitPerPax) = pTotals(DBITEMAirProfitPerPax - 2) / pTotals(DBITEMAirProfitPerPax - 1)
                Else
                    pTotals(DBITEMAirProfitPerPax) = 0
                End If
                If pTotals(DBITEMSerProfitPerPax - 1) <> 0 Then
                    pTotals(DBITEMSerProfitPerPax) = pTotals(DBITEMSerProfitPerPax - 2) / pTotals(DBITEMSerProfitPerPax - 1)
                Else
                    pTotals(DBITEMSerProfitPerPax) = 0
                End If
                If pTotals(DBITEMTotProfitPerPax - 1) <> 0 Then
                    pTotals(DBITEMTotProfitPerPax) = pTotals(DBITEMTotProfitPerPax - 2) / pTotals(DBITEMTotProfitPerPax - 1)
                Else
                    pTotals(DBITEMTotProfitPerPax) = 0
                End If
                .SetCellValue(xlINVCount + 1, 2, "TOTAL")
                For j = 8 To 64
                    .SetCellValue(xlINVCount + 1, j + 1, pTotals(j))
                Next

                For iNeg As Integer = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    For iProf As Integer = DBITEMAirProfitPerPax To DBITEMTotProfitPerPax Step 19
                        If mdsDataSet.Tables(0).Rows(iNeg).Item(iProf) < pTotals(iProf) Then
                            .SetCellStyle(iNeg + 4, iProf + 1, xlStyleNegative)
                        End If
                    Next
                Next

                ' make total row blue
                .SetCellStyle(xlINVCount + 1, 9, xlINVCount + 1, 67, xlStyleHeader)

                ' make profit per pax columns yellow
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Yellow)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General)
                xlStyleHeader.FormatCode = "#,##0.00;-#,##0.00;"
                .SetCellStyle(3, DBITEMAirProfitPerPax + 1, xlINVCount + 1, DBITEMAirProfitPerPax + 1, xlStyleHeader)
                .SetCellStyle(3, DBITEMSerProfitPerPax + 1, xlINVCount + 1, DBITEMSerProfitPerPax + 1, xlStyleHeader)
                .SetCellStyle(3, DBITEMTotProfitPerPax + 1, xlINVCount + 1, DBITEMTotProfitPerPax + 1, xlStyleHeader)

                ' format pax columns without decimals
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(26, xlNumStyle)
                .SetColumnStyle(45, xlNumStyle)
                .SetColumnStyle(64, xlNumStyle)

                .AutoFitColumn(0, 68)

                ' set up grouped columns
                ' for Net Payable details
                .SetCellStyle(3, 9, xlINVCount + 1, 14, xlStyleDetail)
                .SetCellStyle(3, 28, xlINVCount + 1, 33, xlStyleDetail)
                .SetCellStyle(3, 47, xlINVCount + 1, 52, xlStyleDetail)
                .GroupColumns(9, 14)
                .CollapseColumns(15)
                .GroupColumns(28, 33)
                .CollapseColumns(34)
                .GroupColumns(47, 52)
                .CollapseColumns(53)

                ' for IW details
                .SetColumnStyle(17, 23, xlStyleDetail)
                .SetColumnStyle(36, 42, xlStyleDetail)
                .SetColumnStyle(55, 61, xlStyleDetail)
                .GroupColumns(17, 23)
                .CollapseColumns(24)
                .GroupColumns(36, 42)
                .CollapseColumns(43)
                .GroupColumns(55, 61)
                .CollapseColumns(62)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E20_HellasConfidence(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Client Invoices")
                .FreezePanes(1, 0)

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 22, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(16, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, xlNumStyle)
                .SetColumnStyle(14, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(8, xlNumStyle)

                'IssueDate	ClientCode	ClientName	PNR	TicketNumber	Passenger	Details	PaxCount	ProductType	ActionType	InvCode	InvSeries	InvNumber	InvoiceDate	Vessel	BookedBy	Office	ReasonForTravel	CostCentre	RequisitionNumber	NetPayable	TransactionType
                '2019-04-17	010865	Hellas Confidence Shipmanagement S.A.	IGRT	3179724322	PAPAGEORGIOU/PANAGIOTIS MR	ADB-IST-ATH	1	Ticket	Issue	INV		1490576	2019-04-30 00:00:00.000	GIORGOS CONFIDENCE 						335.00	AIR


                .SetCellValue(1, 1, "Issue Date")
                .SetCellValue(1, 2, "Client Code")
                .SetCellValue(1, 3, "Client Name")
                .SetCellValue(1, 4, "PNR")
                .SetCellValue(1, 5, "Ticket Number")
                .SetCellValue(1, 6, "Passenger")
                .SetCellValue(1, 7, "Details")
                .SetCellValue(1, 8, "Pax Count")
                .SetCellValue(1, 9, "Product Type")
                .SetCellValue(1, 10, "Action Type")
                .SetCellValue(1, 11, "Inv Code")
                .SetCellValue(1, 12, "Inv Series")
                .SetCellValue(1, 13, "Inv Number")
                .SetCellValue(1, 14, "Invoice Date")
                .SetCellValue(1, 15, "Vessel")
                .SetCellValue(1, 16, "Net Payable")
                .SetCellValue(1, 17, "Transaction Type")
                .SetCellValue(1, 18, "Booked By")
                .SetCellValue(1, 19, "Office/Dept")
                .SetCellValue(1, 20, "Reason For Travel")
                .SetCellValue(1, 21, "Cost Centre")
                .SetCellValue(1, 22, "Requisition Number")

                .SetCellStyle(1, 1, 1, 22, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    For j = 0 To 21
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(9) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 22, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 22)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E21_ReportByVerifiedUser(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0
        Dim mHours As Decimal = 0

        Try
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Verified Optimize")
                .FreezePanes(1, 0)
                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 20, xlNumStyle)
                xlNumStyle.FormatCode = "dd/MM/yyyy hh:mm"
                .SetColumnStyle(5, xlNumStyle)
                .SetColumnStyle(7, xlNumStyle)
                'xlNumStyle.FormatCode = "HH:mm"
                '.SetColumnStyle(9, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(15, xlNumStyle)
                .SetColumnStyle(16, xlNumStyle)

                'PCC	GDS	PNR	Booked By	Date Logged	Verified By	Date Verified	Verification	HoursDiff	Client Code	Client Name	Client Group	Pax Name	Itinerary	PNR Total	Downsell Total	PNR Fare Basis	Downsell Fare Basis	GDS Pricing	Issuing Carrier						
                'ATHG42100   1A	RVIZCF	VICTORIA DORFANI	26: 17,9	ELENI BLETSOU	20:44,0	ACTIONED	2,9	010649	SHELTON NAVIGATION SA	01 - NP	  1.KISSAMITAKIS/ADRIANOSMR  X 4	15JUL ATH BA/S LHR BA/S EZE	1037,71	697,71	USCYWW	TSCYWW    	FXA/RSEA,U/S8-9	AF						

                .SetCellValue(1, 1, "PCC")
                .SetCellValue(1, 2, "GDS")
                .SetCellValue(1, 3, "PNR")
                .SetCellValue(1, 4, "Booked By")
                .SetCellValue(1, 5, "Date Logged")
                .SetCellValue(1, 6, "Verified By")
                .SetCellValue(1, 7, "Date Verified")
                .SetCellValue(1, 8, "Verification")
                .SetCellValue(1, 9, "Hours Diff")
                .SetCellValue(1, 10, "Client Code")
                .SetCellValue(1, 11, "Client Name")
                .SetCellValue(1, 12, "Client Group")
                .SetCellValue(1, 13, "Pax Name")
                .SetCellValue(1, 14, "Itinerary")
                .SetCellValue(1, 15, "PNR Total")
                .SetCellValue(1, 16, "Downsell Total")
                .SetCellValue(1, 17, "PNR Fare Basis")
                .SetCellValue(1, 18, "Downsell Fare Basis")
                .SetCellValue(1, 19, "GDS Pricing")
                .SetCellValue(1, 20, "Issuing Carrier")

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    For j = 0 To 19
                        If j = 8 Then
                            mHours = mdsDataSet.Tables(0).Rows(i).Item(j)
                            .SetCellValue(xlINVCount, j + 1, Format(Int(mHours), "00") & ":" & Format((mHours - Int(mHours)) * 60, "00"))
                        Else
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        End If

                    Next
                Next

                .AutoFitColumn(1, 20)
            End With
            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub
    Public Sub E22_Euronav(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0
        Dim pSheet As Integer = 0
        Dim ColumnShift As Integer = 0
        If mReport.BooleanOption1 Then
            ColumnShift = 1
        End If
        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "INV")
                pSheet += 1
                Do While pSheet < 3

                    .FreezePanes(1, 0)

                    Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleVoid.Font.FontColor = Color.Gray
                    xlStyleVoid.Font.Italic = True

                    Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleRefund.Font.FontColor = Color.Red

                    Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                    Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleHeader.Font.Bold = True
                    xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                    xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                    xlNumStyle.FormatCode = "@"
                    .SetColumnStyle(1, 18 + ColumnShift, xlNumStyle)
                    xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                    .SetColumnStyle(8, xlNumStyle)
                    xlNumStyle.FormatCode = "dd/mm/yyyy"

                    .SetColumnStyle(3, xlNumStyle)
                    .SetColumnStyle(7, xlNumStyle)
                    .SetCellValue(1, 1, "Invoice Type")
                    .SetCellValue(1, 2, "Invoice Number")
                    .SetCellValue(1, 3, "Invoice Date")
                    .SetCellValue(1, 4, "Vessel Name")
                    .SetCellValue(1, 5, "Routing")
                    .SetCellValue(1, 6, "Destination")
                    .SetCellValue(1, 7, "Departure Date")
                    .SetCellValue(1, 8, "Turnover")
                    If ColumnShift = 1 Then
                        .SetCellValue(1, 9, "Currency")
                    End If
                    .SetCellValue(1, 9 + ColumnShift, "BookedBy")
                    .SetCellValue(1, 10 + ColumnShift, "CPDepartment")
                    .SetCellValue(1, 11 + ColumnShift, "CID Code")
                    .SetCellValue(1, 12 + ColumnShift, "Trip ID")
                    .SetCellValue(1, 13 + ColumnShift, "PassengerName")
                    .SetCellValue(1, 14 + ColumnShift, "Passenger ID")
                    .SetCellValue(1, 15 + ColumnShift, "Nationality")
                    .SetCellValue(1, 16 + ColumnShift, "ClientCode")
                    .SetCellValue(1, 17 + ColumnShift, "PNRID")
                    .SetCellValue(1, 18 + ColumnShift, "ConnectedDocument")
                    .SetCellStyle(1, 1, 1, 18 + ColumnShift, xlStyleHeader)

                    xlINVCount = 1
                    For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                        RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                        If (pSheet = 1 And mdsDataSet.Tables(0).Rows(i).Item(10) <> "Refund") OrElse (pSheet = 2 And mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund") Then
                            If mdsDataSet.Tables(0).Rows(i).Item(3) = "" And mdsDataSet.Tables(0).Rows(i).Item(4) = "" And mdsDataSet.Tables(0).Rows(i).Item(25) = "" Then
                                xlINVCount += 1
                                .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(11) & " " & mdsDataSet.Tables(0).Rows(i).Item(12))
                                .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item(13))
                                .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item(14))
                                .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(15))
                                .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(30))
                                If mdsDataSet.Tables(0).Rows(i).Item(27) = "AIR" And mdsDataSet.Tables(0).Rows(i).Item(30).ToString.Length > 2 Then
                                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(30).ToString.Substring(mdsDataSet.Tables(0).Rows(i).Item(30).ToString.Length - 3))
                                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item(35))
                                Else
                                    .SetCellValue(xlINVCount, 6, "")
                                    .SetCellValue(xlINVCount, 7, "")
                                End If
                                If ColumnShift = 0 Then
                                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(24))
                                Else
                                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(39))
                                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(40))
                                End If
                                .SetCellValue(xlINVCount, 9 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(16))
                                .SetCellValue(xlINVCount, 10 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(17))
                                .SetCellValue(xlINVCount, 11 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(23))
                                .SetCellValue(xlINVCount, 12 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(22))
                                .SetCellValue(xlINVCount, 13 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(7))
                                .SetCellValue(xlINVCount, 14 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(38))
                                .SetCellValue(xlINVCount, 15 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(19))
                                .SetCellValue(xlINVCount, 16 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(1))
                                .SetCellValue(xlINVCount, 17 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(5))
                                .SetCellValue(xlINVCount, 17 + ColumnShift, mdsDataSet.Tables(0).Rows(i).Item(37))
                                If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 18 + ColumnShift, xlStyleOmit)
                                End If
                                If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 18 + ColumnShift, xlStyleVoid)
                                End If
                                If mdsDataSet.Tables(0).Rows(i).Item(11) = "Refund" Then
                                    .SetCellStyle(xlINVCount, 1, xlINVCount, 18 + ColumnShift, xlStyleRefund)
                                End If
                            End If
                        End If
                    Next
                    .AutoFitColumn(1, 18 + ColumnShift)
                    .AddWorksheet("CNS")
                    pSheet += 1
                Loop
                .SelectWorksheet("INV")
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E23_SeaChefs(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim pLastColumn As Integer = 0
        Dim xlINVCount As Integer = 0
        Dim xlHeaderID As Integer = 0
        Dim pInvNumber As String = ""
        Dim pInvRow As Integer = 0
        Dim pInvSum As Decimal = 0
        Dim pBookedBy As New List(Of String)
        Dim pWorksheetCount = 0
        Try

            If mReport.BooleanOption1 Then
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If Not pBookedBy.Contains(mdsDataSet.Tables(0).Rows(i).Item(32)) Then
                        pBookedBy.Add(mdsDataSet.Tables(0).Rows(i).Item(32))
                    End If
                Next
            Else
                pBookedBy.Add("")
            End If

            pBookedBy.Sort()
            For Each bby In pBookedBy
                pLastColumn = 143
                xlINVCount = 0
                xlHeaderID = 0
                pInvNumber = ""
                pInvRow = 0
                pInvSum = 0
                pWorksheetCount += 1
                With xlWorkSheet
                    If pWorksheetCount > 1 Then
                        .AddWorksheet(bby)
                    Else
                        If Not mReport.BooleanOption1 Then
                            .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices")
                        Else
                            .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, bby)
                        End If

                    End If

                    .FreezePanes(8, 0)

                    Dim xlStyleHeaderNotes As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleHeaderNotes.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleHeaderNotes.Fill.SetPatternForegroundColor(Color.FromArgb(255, 169, 208, 142))
                    xlStyleHeaderNotes.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlStyleFixed As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleFixed.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleFixed.Fill.SetPatternForegroundColor(Color.FromArgb(255, 142, 169, 219))
                    xlStyleFixed.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleHeader.Font.Bold = True
                    xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 0, 204, 255))
                    xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlStyleGrey As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleGrey.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleGrey.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                    xlStyleGrey.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                    xlStyleTitle.Font.Bold = True
                    xlStyleTitle.Font.FontSize = 15
                    xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                    xlStyleTitle.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                    xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                    Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                    xlNumStyle.FormatCode = "@"
                    .SetColumnStyle(1, pLastColumn, xlNumStyle)
                    xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                    .SetColumnStyle(10, xlNumStyle)
                    .SetColumnStyle(80, xlNumStyle)
                    xlNumStyle.FormatCode = "dd/mm/yyyy"
                    .SetColumnStyle(11, xlNumStyle)

                    .SetCellValue(3, 3, "Create Invoices")
                    .SetCellStyle(3, 3, xlStyleTitle)

                    .SetCellValue(8, 1, "")
                    .SetCellValue(8, 2, "")
                    .SetCellValue(8, 3, "Changed")
                    .SetCellValue(8, 4, "Row Status")
                    .SetCellValue(8, 5, "*Invoice Header Identifier")
                    .SetCellValue(8, 6, "*Business Unit")
                    .SetCellValue(8, 7, "Import Set")
                    .SetCellValue(8, 8, "*Invoice Number")
                    .SetCellValue(8, 9, "*Invoice Currency")
                    .SetCellValue(8, 10, "*Invoice Amount")
                    .SetCellValue(8, 11, "*Invoice Date")
                    .SetCellValue(8, 12, "**Supplier[..]")
                    .SetCellValue(8, 13, "**Supplier Number")
                    .SetCellValue(8, 14, "*Supplier Site[..]")
                    .SetCellValue(8, 15, "Payment Currency")
                    .SetCellValue(8, 16, "Invoice Type")
                    .SetCellValue(8, 17, "Description")
                    .SetCellValue(8, 18, "Legal Entity Name[..]")
                    .SetCellValue(8, 19, "Payment Terms")
                    .SetCellValue(8, 20, "Terms Date")
                    .SetCellValue(8, 21, "Goods Received Date")
                    .SetCellValue(8, 22, "Invoice Received Date")
                    .SetCellValue(8, 23, "Accounting Date")
                    .SetCellValue(8, 24, "Budget Date")
                    .SetCellValue(8, 25, "Invoice Includes Prepayment")
                    .SetCellValue(8, 26, "Prepayment Number[..]")
                    .SetCellValue(8, 27, "Prepayment Line")
                    .SetCellValue(8, 28, "Prepayment Application Amount")
                    .SetCellValue(8, 29, "Prepayment Accounting Date")
                    .SetCellValue(8, 30, "Payment Method")
                    .SetCellValue(8, 31, "Pay Group")
                    .SetCellValue(8, 32, "Pay Alone")
                    .SetCellValue(8, 33, "Conversion Rate Type")
                    .SetCellValue(8, 34, "Conversion Date")
                    .SetCellValue(8, 35, "Conversion Rate")
                    .SetCellValue(8, 36, "Payment Cross-Conversion Rate Type")
                    .SetCellValue(8, 37, "Payment Cross-Conversion Date")
                    .SetCellValue(8, 38, "Payment Cross-Conversion Rate")
                    .SetCellValue(8, 39, "Discountable Amount")
                    .SetCellValue(8, 40, "Liability Distribution[..]")
                    .SetCellValue(8, 41, "Remit-to Supplier")
                    .SetCellValue(8, 42, "Remit-to Supplier Number")
                    .SetCellValue(8, 43, "Remit-to Address Name")
                    .SetCellValue(8, 44, "Remit-to Account Number")
                    .SetCellValue(8, 45, "Document Category")
                    .SetCellValue(8, 46, "Voucher Number")
                    .SetCellValue(8, 47, "First-Party Tax Registration Number")
                    .SetCellValue(8, 48, "Supplier Tax Registration Number")
                    .SetCellValue(8, 49, "Requester")
                    .SetCellValue(8, 50, "Delivery Channel")
                    .SetCellValue(8, 51, "Bank Charge Bearer")
                    .SetCellValue(8, 52, "Settlement Priority")
                    .SetCellValue(8, 53, "Unique Remittance Identifier")
                    .SetCellValue(8, 54, "Unique Remittance Identifier Check Digit")
                    .SetCellValue(8, 55, "Payment Reason")
                    .SetCellValue(8, 56, "Payment Reason Comments")
                    .SetCellValue(8, 57, "Remittance Message 1")
                    .SetCellValue(8, 58, "Remittance Message 2")
                    .SetCellValue(8, 59, "Remittance Message 3")
                    .SetCellValue(8, 60, "Taxation Country")
                    .SetCellValue(8, 61, "Document Subtype")
                    .SetCellValue(8, 62, "Invoice Internal Sequence")
                    .SetCellValue(8, 63, "Tax Related Invoice")
                    .SetCellValue(8, 64, "Supplier Tax Invoice Number")
                    .SetCellValue(8, 65, "Internal Recording Date")
                    .SetCellValue(8, 66, "Supplier Tax Invoice Date")
                    .SetCellValue(8, 67, "Supplier Tax Invoice Conversion Rate")
                    .SetCellValue(8, 68, "Customs Location Code")
                    .SetCellValue(8, 69, "Correction Year")
                    .SetCellValue(8, 70, "Correction Period")
                    .SetCellValue(8, 71, "Tax Control Amount")
                    .SetCellValue(8, 72, "URL Attachment")
                    .SetCellValue(8, 73, "Context Value[..]")
                    .SetCellValue(8, 74, "Additional Information[..]")
                    .SetCellValue(8, 75, "Regional Context Value [..]")
                    .SetCellValue(8, 76, "Regional Information [..]")
                    .SetCellValue(8, 77, "[..]")
                    .SetCellValue(8, 78, "Line")
                    .SetCellValue(8, 79, "*Type")
                    .SetCellValue(8, 80, "*Amount")
                    .SetCellValue(8, 81, "Invoiced Quantity")
                    .SetCellValue(8, 82, "Unit Price")
                    .SetCellValue(8, 83, "UOM")
                    .SetCellValue(8, 84, "Description")
                    .SetCellValue(8, 85, "Purchase Order[..]")
                    .SetCellValue(8, 86, "Purchase Order Line[..]")
                    .SetCellValue(8, 87, "Purchase Order Schedule[..]")
                    .SetCellValue(8, 88, "Purchase Order Distribution[..]")
                    .SetCellValue(8, 89, "Item Description")
                    .SetCellValue(8, 90, "Receipt[..]")
                    .SetCellValue(8, 91, "Receipt Line[..]")
                    .SetCellValue(8, 92, "Consumption Advice[..]")
                    .SetCellValue(8, 93, "Consumption Advice Line Number[..]")
                    .SetCellValue(8, 94, "Landed Cost Enabled")
                    .SetCellValue(8, 95, "Final Match")
                    .SetCellValue(8, 96, "Distribution Combination[..]")
                    .SetCellValue(8, 97, "Distribution Set[..]")
                    .SetCellValue(8, 98, "Accounting Date")
                    .SetCellValue(8, 99, "Overlay Account Segment")
                    .SetCellValue(8, 100, "Overlay Primary Balancing Segment")
                    .SetCellValue(8, 101, "Overlay Cost Center Segment")
                    .SetCellValue(8, 102, "Budget Date")
                    .SetCellValue(8, 103, "Tax Classification Code")
                    .SetCellValue(8, 104, "Ship-to Location[..]")
                    .SetCellValue(8, 105, "Ship-from Location[..]")
                    .SetCellValue(8, 106, "Location of Final Discharge[..]")
                    .SetCellValue(8, 107, "Regime Code")
                    .SetCellValue(8, 108, "Tax Code")
                    .SetCellValue(8, 109, "Jurisdiction Code")
                    .SetCellValue(8, 110, "Tax Status Code")
                    .SetCellValue(8, 111, "Rate Code")
                    .SetCellValue(8, 112, "Rate")
                    .SetCellValue(8, 113, "Withholding Tax Group")
                    .SetCellValue(8, 114, "Income Tax Type")
                    .SetCellValue(8, 115, "Income Tax Region")
                    .SetCellValue(8, 116, "Prorate Across All Item Lines")
                    .SetCellValue(8, 117, "Line Group Number")
                    .SetCellValue(8, 118, "Transaction Business Category")
                    .SetCellValue(8, 119, "Product Fiscal Classification")
                    .SetCellValue(8, 120, "Intended Use")
                    .SetCellValue(8, 121, "User-Defined Fiscal Classification")
                    .SetCellValue(8, 122, "Product Type")
                    .SetCellValue(8, 123, "Assessable Value")
                    .SetCellValue(8, 124, "Product Category")
                    .SetCellValue(8, 125, "Tax Control Amount")
                    .SetCellValue(8, 126, "Statistical Quantity")
                    .SetCellValue(8, 127, "Deferred Accounting Option")
                    .SetCellValue(8, 128, "Multiperiod Accounting Start Date")
                    .SetCellValue(8, 129, "Multiperiod Accounting End Date")
                    .SetCellValue(8, 130, "Track as Asset")
                    .SetCellValue(8, 131, "Serial Number")
                    .SetCellValue(8, 132, "Book")
                    .SetCellValue(8, 133, "Asset Category")
                    .SetCellValue(8, 134, "Manufacturer")
                    .SetCellValue(8, 135, "Model")
                    .SetCellValue(8, 136, "Requester")
                    .SetCellValue(8, 137, "Item ID")
                    .SetCellValue(8, 138, "Context Value[..]")
                    .SetCellValue(8, 139, "Additional Information[..]")
                    .SetCellValue(8, 140, "Project Information[..]")
                    .SetCellValue(8, 141, "Fiscal Charge Type")
                    .SetCellValue(8, 142, "Multiperiod Accounting Accrual Account[..]")
                    .SetCellValue(8, 143, "Key")


                    .SetCellStyle(8, 1, 8, pLastColumn, xlStyleHeader)
                    .SetRowHeight(8, 56.25)

                    xlINVCount = 8
                    pInvNumber = mdsDataSet.Tables(0).Rows(0).Item(4)
                    pInvRow = xlINVCount + 1
                    pInvSum = 0
                    xlHeaderID = 1
                    For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                        If Not mReport.BooleanOption1 Or mdsDataSet.Tables(0).Rows(i).Item(32) = bby Then
                            RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                            xlINVCount += 1
                            If mdsDataSet.Tables(0).Rows(i).Item(4) = pInvNumber Then
                                pInvSum += mdsDataSet.Tables(0).Rows(i).Item(16)
                            Else
                                xlHeaderID += 1
                                .SetCellValue(pInvRow, 10, pInvSum)
                                For k = pInvRow + 1 To xlINVCount - 1
                                    .SetCellStyle(k, 6, xlStyleGrey)
                                    .SetCellStyle(k, 10, xlStyleGrey)
                                    .SetCellStyle(k, 47, xlStyleGrey)
                                    .SetCellValue(k, 6, "")
                                    .SetCellValue(k, 10, "")
                                    .SetCellValue(k, 47, "")
                                Next
                                pInvSum = mdsDataSet.Tables(0).Rows(i).Item(16)
                                pInvNumber = mdsDataSet.Tables(0).Rows(i).Item(4)
                                pInvRow = xlINVCount
                            End If
                            ' mdsDataSet.Tables(0).Rows(i).Item(32) is booked by
                            .SetCellValue(xlINVCount, 5, xlHeaderID)
                            .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(3))
                            .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(4))
                            .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(5))

                            .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(7))
                            .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item(8))
                            .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(9))
                            .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item(10))
                            .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item(11))
                            .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item(12))
                            .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(i).Item(13))

                            .SetCellValue(xlINVCount, 78, mdsDataSet.Tables(0).Rows(i).Item(14))
                            .SetCellValue(xlINVCount, 79, mdsDataSet.Tables(0).Rows(i).Item(15))
                            .SetCellValue(xlINVCount, 80, mdsDataSet.Tables(0).Rows(i).Item(16))

                            .SetCellValue(xlINVCount, 84, mdsDataSet.Tables(0).Rows(i).Item(17))
                            .SetCellValue(xlINVCount, 96, mdsDataSet.Tables(0).Rows(i).Item(19))
                            .SetCellValue(xlINVCount, 103, mdsDataSet.Tables(0).Rows(i).Item(44))
                        End If
                    Next
                    If pInvNumber <> "" AndAlso pInvSum <> 0 Then
                        .SetCellValue(pInvRow, 10, pInvSum)
                        For k = pInvRow + 1 To xlINVCount
                            .SetCellStyle(k, 6, xlStyleGrey)
                            .SetCellStyle(k, 10, xlStyleGrey)
                            .SetCellStyle(k, 47, xlStyleGrey)
                            .SetCellValue(k, 6, "")
                            .SetCellValue(k, 10, "")
                            .SetCellValue(k, 47, "")
                        Next
                    End If
                    .AutoFitColumn(1, pLastColumn)
                    .HideColumn(7)
                    .HideColumn(18, 46)
                    .HideColumn(48, 77)
                    .HideColumn(81, 83)
                    .HideColumn(85, 95)
                    .HideColumn(97, 102)
                    .HideColumn(104, 143)

                End With
            Next
            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E23_SeaChefsxx(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim pLastColumn As Integer = 0
        Dim xlINVCount As Integer = 0
        Dim xlHeaderID As Integer = 0
        Dim pInvNumber As String = ""
        Dim pInvRow As Integer = 0
        Dim pInvSum As Decimal = 0

        Try
            pLastColumn = 143
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices")

                .FreezePanes(8, 0)

                Dim xlStyleHeaderNotes As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeaderNotes.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeaderNotes.Fill.SetPatternForegroundColor(Color.FromArgb(255, 169, 208, 142))
                xlStyleHeaderNotes.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleFixed As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleFixed.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleFixed.Fill.SetPatternForegroundColor(Color.FromArgb(255, 142, 169, 219))
                xlStyleFixed.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 0, 204, 255))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleGrey As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGrey.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGrey.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleGrey.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTitle.Font.Bold = True
                xlStyleTitle.Font.FontSize = 15
                xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleTitle.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, pLastColumn, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(10, xlNumStyle)
                .SetColumnStyle(80, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(11, xlNumStyle)

                .SetCellValue(3, 3, "Create Invoices")
                .SetCellStyle(3, 3, xlStyleTitle)

                .SetCellValue(8, 1, "")
                .SetCellValue(8, 2, "")
                .SetCellValue(8, 3, "Changed")
                .SetCellValue(8, 4, "Row Status")
                .SetCellValue(8, 5, "*Invoice Header Identifier")
                .SetCellValue(8, 6, "*Business Unit")
                .SetCellValue(8, 7, "Import Set")
                .SetCellValue(8, 8, "*Invoice Number")
                .SetCellValue(8, 9, "*Invoice Currency")
                .SetCellValue(8, 10, "*Invoice Amount")
                .SetCellValue(8, 11, "*Invoice Date")
                .SetCellValue(8, 12, "**Supplier[..]")
                .SetCellValue(8, 13, "**Supplier Number")
                .SetCellValue(8, 14, "*Supplier Site[..]")
                .SetCellValue(8, 15, "Payment Currency")
                .SetCellValue(8, 16, "Invoice Type")
                .SetCellValue(8, 17, "Description")
                .SetCellValue(8, 18, "Legal Entity Name[..]")
                .SetCellValue(8, 19, "Payment Terms")
                .SetCellValue(8, 20, "Terms Date")
                .SetCellValue(8, 21, "Goods Received Date")
                .SetCellValue(8, 22, "Invoice Received Date")
                .SetCellValue(8, 23, "Accounting Date")
                .SetCellValue(8, 24, "Budget Date")
                .SetCellValue(8, 25, "Invoice Includes Prepayment")
                .SetCellValue(8, 26, "Prepayment Number[..]")
                .SetCellValue(8, 27, "Prepayment Line")
                .SetCellValue(8, 28, "Prepayment Application Amount")
                .SetCellValue(8, 29, "Prepayment Accounting Date")
                .SetCellValue(8, 30, "Payment Method")
                .SetCellValue(8, 31, "Pay Group")
                .SetCellValue(8, 32, "Pay Alone")
                .SetCellValue(8, 33, "Conversion Rate Type")
                .SetCellValue(8, 34, "Conversion Date")
                .SetCellValue(8, 35, "Conversion Rate")
                .SetCellValue(8, 36, "Payment Cross-Conversion Rate Type")
                .SetCellValue(8, 37, "Payment Cross-Conversion Date")
                .SetCellValue(8, 38, "Payment Cross-Conversion Rate")
                .SetCellValue(8, 39, "Discountable Amount")
                .SetCellValue(8, 40, "Liability Distribution[..]")
                .SetCellValue(8, 41, "Remit-to Supplier")
                .SetCellValue(8, 42, "Remit-to Supplier Number")
                .SetCellValue(8, 43, "Remit-to Address Name")
                .SetCellValue(8, 44, "Remit-to Account Number")
                .SetCellValue(8, 45, "Document Category")
                .SetCellValue(8, 46, "Voucher Number")
                .SetCellValue(8, 47, "First-Party Tax Registration Number")
                .SetCellValue(8, 48, "Supplier Tax Registration Number")
                .SetCellValue(8, 49, "Requester")
                .SetCellValue(8, 50, "Delivery Channel")
                .SetCellValue(8, 51, "Bank Charge Bearer")
                .SetCellValue(8, 52, "Settlement Priority")
                .SetCellValue(8, 53, "Unique Remittance Identifier")
                .SetCellValue(8, 54, "Unique Remittance Identifier Check Digit")
                .SetCellValue(8, 55, "Payment Reason")
                .SetCellValue(8, 56, "Payment Reason Comments")
                .SetCellValue(8, 57, "Remittance Message 1")
                .SetCellValue(8, 58, "Remittance Message 2")
                .SetCellValue(8, 59, "Remittance Message 3")
                .SetCellValue(8, 60, "Taxation Country")
                .SetCellValue(8, 61, "Document Subtype")
                .SetCellValue(8, 62, "Invoice Internal Sequence")
                .SetCellValue(8, 63, "Tax Related Invoice")
                .SetCellValue(8, 64, "Supplier Tax Invoice Number")
                .SetCellValue(8, 65, "Internal Recording Date")
                .SetCellValue(8, 66, "Supplier Tax Invoice Date")
                .SetCellValue(8, 67, "Supplier Tax Invoice Conversion Rate")
                .SetCellValue(8, 68, "Customs Location Code")
                .SetCellValue(8, 69, "Correction Year")
                .SetCellValue(8, 70, "Correction Period")
                .SetCellValue(8, 71, "Tax Control Amount")
                .SetCellValue(8, 72, "URL Attachment")
                .SetCellValue(8, 73, "Context Value[..]")
                .SetCellValue(8, 74, "Additional Information[..]")
                .SetCellValue(8, 75, "Regional Context Value [..]")
                .SetCellValue(8, 76, "Regional Information [..]")
                .SetCellValue(8, 77, "[..]")
                .SetCellValue(8, 78, "Line")
                .SetCellValue(8, 79, "*Type")
                .SetCellValue(8, 80, "*Amount")
                .SetCellValue(8, 81, "Invoiced Quantity")
                .SetCellValue(8, 82, "Unit Price")
                .SetCellValue(8, 83, "UOM")
                .SetCellValue(8, 84, "Description")
                .SetCellValue(8, 85, "Purchase Order[..]")
                .SetCellValue(8, 86, "Purchase Order Line[..]")
                .SetCellValue(8, 87, "Purchase Order Schedule[..]")
                .SetCellValue(8, 88, "Purchase Order Distribution[..]")
                .SetCellValue(8, 89, "Item Description")
                .SetCellValue(8, 90, "Receipt[..]")
                .SetCellValue(8, 91, "Receipt Line[..]")
                .SetCellValue(8, 92, "Consumption Advice[..]")
                .SetCellValue(8, 93, "Consumption Advice Line Number[..]")
                .SetCellValue(8, 94, "Landed Cost Enabled")
                .SetCellValue(8, 95, "Final Match")
                .SetCellValue(8, 96, "Distribution Combination[..]")
                .SetCellValue(8, 97, "Distribution Set[..]")
                .SetCellValue(8, 98, "Accounting Date")
                .SetCellValue(8, 99, "Overlay Account Segment")
                .SetCellValue(8, 100, "Overlay Primary Balancing Segment")
                .SetCellValue(8, 101, "Overlay Cost Center Segment")
                .SetCellValue(8, 102, "Budget Date")
                .SetCellValue(8, 103, "Tax Classification Code")
                .SetCellValue(8, 104, "Ship-to Location[..]")
                .SetCellValue(8, 105, "Ship-from Location[..]")
                .SetCellValue(8, 106, "Location of Final Discharge[..]")
                .SetCellValue(8, 107, "Regime Code")
                .SetCellValue(8, 108, "Tax Code")
                .SetCellValue(8, 109, "Jurisdiction Code")
                .SetCellValue(8, 110, "Tax Status Code")
                .SetCellValue(8, 111, "Rate Code")
                .SetCellValue(8, 112, "Rate")
                .SetCellValue(8, 113, "Withholding Tax Group")
                .SetCellValue(8, 114, "Income Tax Type")
                .SetCellValue(8, 115, "Income Tax Region")
                .SetCellValue(8, 116, "Prorate Across All Item Lines")
                .SetCellValue(8, 117, "Line Group Number")
                .SetCellValue(8, 118, "Transaction Business Category")
                .SetCellValue(8, 119, "Product Fiscal Classification")
                .SetCellValue(8, 120, "Intended Use")
                .SetCellValue(8, 121, "User-Defined Fiscal Classification")
                .SetCellValue(8, 122, "Product Type")
                .SetCellValue(8, 123, "Assessable Value")
                .SetCellValue(8, 124, "Product Category")
                .SetCellValue(8, 125, "Tax Control Amount")
                .SetCellValue(8, 126, "Statistical Quantity")
                .SetCellValue(8, 127, "Deferred Accounting Option")
                .SetCellValue(8, 128, "Multiperiod Accounting Start Date")
                .SetCellValue(8, 129, "Multiperiod Accounting End Date")
                .SetCellValue(8, 130, "Track as Asset")
                .SetCellValue(8, 131, "Serial Number")
                .SetCellValue(8, 132, "Book")
                .SetCellValue(8, 133, "Asset Category")
                .SetCellValue(8, 134, "Manufacturer")
                .SetCellValue(8, 135, "Model")
                .SetCellValue(8, 136, "Requester")
                .SetCellValue(8, 137, "Item ID")
                .SetCellValue(8, 138, "Context Value[..]")
                .SetCellValue(8, 139, "Additional Information[..]")
                .SetCellValue(8, 140, "Project Information[..]")
                .SetCellValue(8, 141, "Fiscal Charge Type")
                .SetCellValue(8, 142, "Multiperiod Accounting Accrual Account[..]")
                .SetCellValue(8, 143, "Key")


                .SetCellStyle(8, 1, 8, pLastColumn, xlStyleHeader)
                .SetRowHeight(8, 56.25)

                xlINVCount = 8
                pInvNumber = mdsDataSet.Tables(0).Rows(0).Item(4)
                pInvRow = xlINVCount + 1
                pInvSum = 0
                xlHeaderID = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    If mdsDataSet.Tables(0).Rows(i).Item(4) = pInvNumber Then
                        pInvSum += mdsDataSet.Tables(0).Rows(i).Item(16)
                    Else
                        xlHeaderID += 1
                        .SetCellValue(pInvRow, 10, pInvSum)
                        For k = pInvRow + 1 To xlINVCount - 1
                            .SetCellStyle(k, 6, xlStyleGrey)
                            .SetCellStyle(k, 10, xlStyleGrey)
                            .SetCellStyle(k, 47, xlStyleGrey)
                            .SetCellValue(k, 6, "")
                            .SetCellValue(k, 10, "")
                            .SetCellValue(k, 47, "")
                        Next
                        pInvSum = mdsDataSet.Tables(0).Rows(i).Item(16)
                        pInvNumber = mdsDataSet.Tables(0).Rows(i).Item(4)
                        pInvRow = xlINVCount
                    End If
                    ' mdsDataSet.Tables(0).Rows(i).Item(32) is booked by
                    .SetCellValue(xlINVCount, 5, xlHeaderID)
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(5))

                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(xlINVCount, 47, mdsDataSet.Tables(0).Rows(i).Item(13))

                    .SetCellValue(xlINVCount, 78, mdsDataSet.Tables(0).Rows(i).Item(14))
                    .SetCellValue(xlINVCount, 79, mdsDataSet.Tables(0).Rows(i).Item(15))
                    .SetCellValue(xlINVCount, 80, mdsDataSet.Tables(0).Rows(i).Item(16))

                    .SetCellValue(xlINVCount, 84, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(xlINVCount, 96, mdsDataSet.Tables(0).Rows(i).Item(19))
                    .SetCellValue(xlINVCount, 103, mdsDataSet.Tables(0).Rows(i).Item(44))

                Next
                If pInvNumber <> "" AndAlso pInvSum <> 0 Then
                    .SetCellValue(pInvRow, 10, pInvSum)
                    For k = pInvRow + 1 To xlINVCount
                        .SetCellStyle(k, 6, xlStyleGrey)
                        .SetCellStyle(k, 10, xlStyleGrey)
                        .SetCellStyle(k, 47, xlStyleGrey)
                        .SetCellValue(k, 6, "")
                        .SetCellValue(k, 10, "")
                        .SetCellValue(k, 47, "")
                    Next
                End If
                .AutoFitColumn(1, pLastColumn)
                .HideColumn(7)
                .HideColumn(18, 46)
                .HideColumn(48, 77)
                .HideColumn(81, 83)
                .HideColumn(85, 95)
                .HideColumn(97, 102)
                .HideColumn(104, 143)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E29_SeaChefsDetailed(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim pLastColumn As Integer = 0
        Dim xlINVCount As Integer = 0
        Dim xlHeaderID As Integer = 0
        Dim pInvNumber As String = ""
        Dim pInvRow As Integer = 0
        Dim pInvSum As Decimal = 0

        Try
            pLastColumn = 23
            If mReport.BooleanOption1 Then
                pLastColumn = 44
            End If
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices")

                .FreezePanes(8, 0)

                Dim xlStyleHeaderNotes As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeaderNotes.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeaderNotes.Fill.SetPatternForegroundColor(Color.FromArgb(255, 169, 208, 142))
                xlStyleHeaderNotes.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleFixed As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleFixed.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleFixed.Fill.SetPatternForegroundColor(Color.FromArgb(255, 142, 169, 219))
                xlStyleFixed.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 0, 204, 255))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleGrey As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGrey.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGrey.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleGrey.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTitle.Font.Bold = True
                xlStyleTitle.Font.FontSize = 15
                xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleTitle.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, pLastColumn, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(7, xlNumStyle)
                .SetColumnStyle(17, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(8, xlNumStyle)

                .SetCellValue(3, 3, "Create Invoices")
                .SetCellStyle(3, 3, xlStyleTitle)

                .SetCellValue(7, 1, "")
                .SetCellValue(7, 2, "")
                .SetCellValue(7, 3, "")
                .SetCellValue(7, 4, "FIXED")
                .SetCellValue(7, 5, "Invoice")
                .SetCellValue(7, 6, "FIXED")
                .SetCellValue(7, 7, "")
                .SetCellValue(7, 8, "Batch Date")
                .SetCellValue(7, 9, "FIXED")
                .SetCellValue(7, 10, "FIXED")
                .SetCellValue(7, 11, "FIXED")
                .SetCellValue(7, 12, "FIXED")
                .SetCellValue(7, 13, "Standard/Credit Memo")
                .SetCellValue(7, 14, "FIXED")
                .SetCellValue(7, 15, "FIXED")
                .SetCellValue(7, 16, "FIXED")
                .SetCellValue(7, 17, "Payable")
                .SetCellValue(7, 18, "AL / Booked By / Routing / Traveller / Dep date/TRID")
                .SetCellValue(7, 19, "Flight / P & I / Accommodation")
                .SetCellValue(7, 20, "Cost Centre")
                .SetCellValue(7, 21, "Vessel")
                .SetCellValue(7, 22, "Client Code")
                .SetCellValue(7, 23, "Client Name")

                .SetCellStyle(7, 4, 7, 4, xlStyleFixed)
                .SetCellStyle(7, 6, 7, 6, xlStyleFixed)
                .SetCellStyle(7, 9, 7, 12, xlStyleFixed)
                .SetCellStyle(7, 14, 7, 16, xlStyleFixed)
                .SetCellStyle(7, 5, 7, 5, xlStyleHeaderNotes)
                .SetCellStyle(7, 8, 7, 8, xlStyleHeaderNotes)
                .SetCellStyle(7, 13, 7, 13, xlStyleHeaderNotes)
                .SetCellStyle(7, 17, 7, pLastColumn, xlStyleHeaderNotes)

                .SetCellValue(8, 1, "Changed")
                .SetCellValue(8, 2, "Row Status")
                .SetCellValue(8, 3, "Invoice Header Identifier")
                .SetCellValue(8, 4, "Business Unit")
                .SetCellValue(8, 5, "Invoice Number")
                .SetCellValue(8, 6, "Invoice Currency")
                .SetCellValue(8, 7, "Invoice Amount")
                .SetCellValue(8, 8, "Invoice Date")
                .SetCellValue(8, 9, "Supplier")
                .SetCellValue(8, 10, "Supplier Number")
                .SetCellValue(8, 11, "Supplier Site")
                .SetCellValue(8, 12, "Payment Currency")
                .SetCellValue(8, 13, "Invoice Type")
                .SetCellValue(8, 14, "First-Party Tax Registration Number")
                .SetCellValue(8, 15, "Line")
                .SetCellValue(8, 16, "Type")
                .SetCellValue(8, 17, "Amount")
                .SetCellValue(8, 18, "Description")
                .SetCellValue(8, 19, "Expense Type")
                .SetCellValue(8, 20, "Distribution Combination")
                .SetCellValue(8, 21, "Vessel")
                .SetCellValue(8, 22, "Client Code")
                .SetCellValue(8, 23, "Client Name")

                If mReport.BooleanOption1 Then
                    .SetCellValue(8, 24, "Invoice Code")
                    .SetCellValue(8, 25, "Invoice Series")
                    .SetCellValue(8, 26, "PNR")
                    .SetCellValue(8, 27, "Airline Code")
                    .SetCellValue(8, 28, "Ticket Number")
                    .SetCellValue(8, 29, "Pax Count")
                    .SetCellValue(8, 30, "Product Type")
                    .SetCellValue(8, 31, "Product Type LT")
                    .SetCellValue(8, 32, "Action Type")
                    .SetCellValue(8, 33, "Id")
                    .SetCellValue(8, 34, "Office")
                    .SetCellValue(8, 35, "Reason For Travel")
                    .SetCellValue(8, 36, "Requisition Number")
                    .SetCellValue(8, 37, "OPT")
                    .SetCellValue(8, 38, "TRID")
                    .SetCellValue(8, 39, "Verified")
                    .SetCellValue(8, 40, "Remarks")
                    .SetCellValue(8, 41, "RegNr")
                    .SetCellValue(8, 42, "Sales Person")
                    .SetCellValue(8, 43, "Issuing Agent")
                    .SetCellValue(8, 44, "Creator Agent")
                End If

                .SetCellStyle(8, 1, 8, pLastColumn, xlStyleHeader)
                .SetRowHeight(8, 56.25)

                xlINVCount = 8
                pInvNumber = mdsDataSet.Tables(0).Rows(0).Item(4)
                pInvRow = xlINVCount + 1
                pInvSum = 0
                xlHeaderID = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    If mdsDataSet.Tables(0).Rows(i).Item(4) = pInvNumber Then
                        pInvSum += mdsDataSet.Tables(0).Rows(i).Item(16)
                    Else
                        xlHeaderID += 1
                        .SetCellValue(pInvRow, 7, pInvSum)
                        For k = pInvRow + 1 To xlINVCount - 1
                            .SetCellStyle(k, 4, xlStyleGrey)
                            .SetCellStyle(k, 7, xlStyleGrey)
                            .SetCellStyle(k, 14, xlStyleGrey)
                            .SetCellValue(k, 4, "")
                            .SetCellValue(k, 7, "")
                            .SetCellValue(k, 14, "")
                        Next
                        pInvSum = mdsDataSet.Tables(0).Rows(i).Item(16)
                        pInvNumber = mdsDataSet.Tables(0).Rows(i).Item(4)
                        pInvRow = xlINVCount
                    End If
                    .SetCellValue(xlINVCount, 3, xlHeaderID)
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(5))

                    For j = 8 To pLastColumn
                        .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j - 1))
                    Next
                    If mReport.BooleanOption1 Then
                        For j = pLastColumn + 1 To pLastColumn
                            .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j - 1))
                        Next

                    End If
                Next
                If pInvNumber <> "" AndAlso pInvSum <> 0 Then
                    .SetCellValue(pInvRow, 7, pInvSum)
                    For k = pInvRow + 1 To xlINVCount
                        .SetCellStyle(k, 4, xlStyleGrey)
                        .SetCellStyle(k, 7, xlStyleGrey)
                        .SetCellStyle(k, 14, xlStyleGrey)
                        .SetCellValue(k, 4, "")
                        .SetCellValue(k, 7, "")
                        .SetCellValue(k, 14, "")
                    Next
                End If
                .AutoFitColumn(1, pLastColumn)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E24_ProfitPerAgentTotals(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet

                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 10, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(6, 8, xlNumStyle)
                .SetColumnStyle(10, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(9, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                '                .SetColumnStyle(3, xlNumStyle)

                .SetCellValue(1, 1, "Group Name")
                .SetCellValue(1, 2, "Creator Agent")
                .SetCellValue(1, 3, "Sales Person")
                .SetCellValue(1, 4, "Client Code")
                .SetCellValue(1, 5, "Client Name")
                .SetCellValue(1, 6, "Sales")
                .SetCellValue(1, 7, "Cost")
                .SetCellValue(1, 8, "Profit")
                .SetCellValue(1, 9, "Pax")
                .SetCellValue(1, 10, "Profit per Pax")

                .SetCellStyle(1, 1, 1, 10, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 1 To 10
                        .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j - 1))
                    Next
                Next
                .AutoFitColumn(1, 10)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E25_ProfitPerAgentTransactions(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet

                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 12, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(8, 10, xlNumStyle)
                .SetColumnStyle(12, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(11, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(6, xlNumStyle)

                .SetCellValue(1, 1, "Group Name")
                .SetCellValue(1, 2, "Creator Agent")
                .SetCellValue(1, 3, "Sales Person")
                .SetCellValue(1, 4, "Client Code")
                .SetCellValue(1, 5, "Client Name")
                .SetCellValue(1, 6, "Issue Date")
                .SetCellValue(1, 7, "Doc Number")
                .SetCellValue(1, 8, "Sales")
                .SetCellValue(1, 9, "Cost")
                .SetCellValue(1, 10, "Profit")
                .SetCellValue(1, 11, "Pax")
                .SetCellValue(1, 12, "Profit per Pax")

                .SetCellStyle(1, 1, 1, 10, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 1 To 12
                        .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j - 1))
                    Next
                Next
                .AutoFitColumn(1, 12)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E28_OptimizationSavings(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        ThisDocument = New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try

            With ThisDocument

                .FreezePanes(1, 0)

                Dim xlStyleCurrency As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCurrency.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleCurrency.Fill.SetPatternForegroundColor(Color.Pink)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 28, xlNumStyle)
                .SetColumnStyle(11, xlDecStyle)
                .SetColumnStyle(13, 14, xlDecStyle)
                .SetColumnStyle(16, xlDecStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(7, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0"
                .SetColumnStyle(10, xlNumStyle)
                .SetColumnStyle(22, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy hh:MM"
                .SetColumnStyle(20, 21, xlNumStyle)
                .SetCellValue(1, 1, "Record Locator")
                .SetCellValue(1, 2, "Pseudocity")
                .SetCellValue(1, 3, "Consultant")
                .SetCellValue(1, 4, "Consultant Group")
                .SetCellValue(1, 5, "Account Code")
                .SetCellValue(1, 6, "Account Name")
                .SetCellValue(1, 7, "Date Of Travel")
                .SetCellValue(1, 8, "IATA Itinerary")
                .SetCellValue(1, 9, "Plate Carrier")
                .SetCellValue(1, 10, "No of Pax")
                .SetCellValue(1, 11, "Fare Amount")
                .SetCellValue(1, 12, "Fare Currency")
                .SetCellValue(1, 13, "Potential Downsell")
                .SetCellValue(1, 14, "Potential Savings")
                .SetCellValue(1, 15, "Downsell Currency")
                .SetCellValue(1, 16, "Actual Savings")
                .SetCellValue(1, 17, "Verification")
                .SetCellValue(1, 18, "Actioned By")
                .SetCellValue(1, 19, "Actioned Group")
                .SetCellValue(1, 20, "Date Logged")
                .SetCellValue(1, 21, "Date Verified")
                .SetCellValue(1, 22, "Delay (Minutes)")
                .SetCellValue(1, 23, "Delay (Hours)")
                .SetCellValue(1, 24, "Action Reason")
                .SetCellValue(1, 25, "Agent Sign")
                .SetCellValue(1, 26, "Original Fare Basis")
                .SetCellValue(1, 27, "New Fare Basis")
                .SetCellValue(1, 28, "Pricing Command")
                .SetCellValue(1, 29, "Client Group")

                .SetCellStyle(1, 1, 1, 29, xlStyleHeader)

                Dim VerificationList As New Reports.E28_VerificationList
                Dim ActionReasonList As New Reports.E28_ActionReasonList
                Dim ConsultantsList As New Reports.E28_GroupSubItemList
                Dim ActionedByList As New Reports.E28_GroupSubItemList
                Dim ClientList As New Reports.E28_GroupSubItemList
                Dim HourLogged As Integer
                Dim NoOfPax As Integer
                Dim FareAmount As Decimal
                Dim DownsellTotal As Decimal
                Dim ClientGroup As String
                Dim AccountCode As String
                Dim AccountName As String
                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    ClientGroup = mdsDataSet.Tables(0).Rows(i).Item("ClientGroup")
                    AccountCode = mdsDataSet.Tables(0).Rows(i).Item("AccountCode")
                    AccountName = mdsDataSet.Tables(0).Rows(i).Item("AccountName")
                    If ClientGroup = "" Then ClientGroup = "[No Group]"
                    If AccountCode = "" Then AccountCode = "[No Code]"
                    If AccountName = "" Then AccountName = "[No Name]"
                    HourLogged = Convert.ToDateTime(mdsDataSet.Tables(0).Rows(i).Item("DateLogged")).Hour
                    NoOfPax = mdsDataSet.Tables(0).Rows(i).Item("NoOfPax")
                    FareAmount = mdsDataSet.Tables(0).Rows(i).Item("FareAmount") * NoOfPax
                    DownsellTotal = mdsDataSet.Tables(0).Rows(i).Item("DownsellTotal") * NoOfPax
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item("RecordLocator")) ' "Record Locator")
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item("PseudoCity")) ' "Pseudocity")
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item("GDSUserName")) ' "Consultant")
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item("GDSUserGroup")) ' "Consultant")
                    .SetCellValue(xlINVCount, 5, AccountCode)
                    .SetCellValue(xlINVCount, 6, AccountName)
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item("DateOfTravel")) ' "Date Of Travel")
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item("IataItinerary")) ' "IATA Itinerary")
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item("PlateCarrier")) ' "Plate Carrier")
                    .SetCellValue(xlINVCount, 10, NoOfPax) ' "No of Pax")
                    .SetCellValue(xlINVCount, 11, FareAmount) ' "Fare Amount")
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item("FareCurrency"))
                    .SetCellValue(xlINVCount, 13, DownsellTotal) ' "Potential Downsell")
                    .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                    If mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") <> mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency") Then
                        .SetCellStyle(xlINVCount, 11, xlINVCount, 15, xlStyleCurrency)
                        .SetCellValue(xlINVCount, 14, "")
                        .SetCellValue(xlINVCount, 16, "")
                    Else
                        .SetCellValue(xlINVCount, 14, FareAmount - DownsellTotal) ' "Potential Savings")
                        .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR")) ' "Actual Savings")
                    End If
                    .SetCellValue(xlINVCount, 17, mdsDataSet.Tables(0).Rows(i).Item("VerificationReason")) ' "Verification")
                    .SetCellValue(xlINVCount, 18, mdsDataSet.Tables(0).Rows(i).Item("ActionedBy")) ' "Actioned By")
                    .SetCellValue(xlINVCount, 19, mdsDataSet.Tables(0).Rows(i).Item("ActionedByGroup")) ' "Ops Group")
                    .SetCellValue(xlINVCount, 20, mdsDataSet.Tables(0).Rows(i).Item("DateLogged")) ' "Date Logged")
                    .SetCellValue(xlINVCount, 21, mdsDataSet.Tables(0).Rows(i).Item("VerifiedDate")) ' "Date Verified")
                    .SetCellValue(xlINVCount, 22, mdsDataSet.Tables(0).Rows(i).Item("DiffTime")) ' "Delay (Minutes)")
                    .SetCellValue(xlINVCount, 23, $"{ mdsDataSet.Tables(0).Rows(i).Item("DiffHours").ToString.PadLeft(2, "0")}:{mdsDataSet.Tables(0).Rows(i).Item("DiffMinutes").ToString.PadLeft(2, "0")}") ' "Delay (Minutes)")
                    .SetCellValue(xlINVCount, 24, mdsDataSet.Tables(0).Rows(i).Item("VerificationReasonText")) ' "Action Reason")
                    .SetCellValue(xlINVCount, 25, mdsDataSet.Tables(0).Rows(i).Item("AgentSign")) ' "Amadeus Sign")
                    .SetCellValue(xlINVCount, 26, mdsDataSet.Tables(0).Rows(i).Item("OriginalFareBasis")) ' "Original Fare Basis")
                    .SetCellValue(xlINVCount, 27, mdsDataSet.Tables(0).Rows(i).Item("NewFareBasis")) ' "New Fare Basis")
                    .SetCellValue(xlINVCount, 28, mdsDataSet.Tables(0).Rows(i).Item("PricingCommand")) ' "Pricing Command")
                    .SetCellValue(xlINVCount, 29, ClientGroup)
                    VerificationList.Add(mdsDataSet.Tables(0).Rows(i).Item("VerificationReason"),
                                         mdsDataSet.Tables(0).Rows(i).Item("DiffTime"),
                                         mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR"),
                                         FareAmount - DownsellTotal,
                                         mdsDataSet.Tables(0).Rows(i).Item("FareCurrency"),
                                         mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                    If mdsDataSet.Tables(0).Rows(i).Item("VerificationReason") = E28_ReasonACTIONED Then
                        ConsultantsList.AddActioned(mdsDataSet.Tables(0).Rows(i).Item("GDSUserGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("GDSUserName") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                        ActionedByList.AddActioned(mdsDataSet.Tables(0).Rows(i).Item("ActionedByGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("ActionedBy") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                        ClientList.AddActioned(ClientGroup _
                                                     , $"{AccountCode} {AccountName}", mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                    ElseIf mdsDataSet.Tables(0).Rows(i).Item("VerificationReason") = E28_ReasonPOSTPONE Then
                        ActionReasonList.AddPostponed(mdsDataSet.Tables(0).Rows(i).Item("VerificationReasonText"), HourLogged, mdsDataSet.Tables(0).Rows(i).Item("DiffTime"))
                        ConsultantsList.AddPostponed(mdsDataSet.Tables(0).Rows(i).Item("GDSUserGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("GDSUserName") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                        ActionedByList.AddPostponed(mdsDataSet.Tables(0).Rows(i).Item("ActionedByGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("ActionedBy") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                    ElseIf mdsDataSet.Tables(0).Rows(i).Item("VerificationReason") = E28_ReasonREJECT Then
                        ActionReasonList.AddRejected(mdsDataSet.Tables(0).Rows(i).Item("VerificationReasonText"), HourLogged, mdsDataSet.Tables(0).Rows(i).Item("DiffTime"))
                        ConsultantsList.AddRejected(mdsDataSet.Tables(0).Rows(i).Item("GDSUserGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("GDSUserName") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                        ActionedByList.AddRejected(mdsDataSet.Tables(0).Rows(i).Item("ActionedByGroup") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("ActionedBy") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellAmountPerPNR") _
                                                     , FareAmount - DownsellTotal _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DiffTime") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("FareCurrency") _
                                                     , mdsDataSet.Tables(0).Rows(i).Item("DownsellCurrency"))
                    End If
                Next
                .Filter(1, 1, xlINVCount, 29)
                .AutoFitColumn(1, 29)

                E28_AddActionedOnlyWorksheet("ByClient", ClientList)

                E28_AddVerificationWorksheet("PNR Creator by Group", ConsultantsList, True, True)
                E28_AddVerificationWorksheet("Actioned By by Group", ActionedByList, True, True)

                E28_AddVerificationWorksheet("PNR Creators", ConsultantsList, False, True)
                E28_AddVerificationWorksheet("Actioned By", ActionedByList, False, True)

                E28_AddVerificationWorksheet("PNR Creator Groups", ConsultantsList, True, False)
                E28_AddVerificationWorksheet("Actioned By Groups", ActionedByList, True, False)

                RowCount = 0
                .AddWorksheet("Action Totals")
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 6, xlNumStyle)
                .SetColumnStyle(4, 6, xlDecStyle)
                xlNumStyle.FormatCode = "#,##0"
                .SetColumnStyle(2, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00%"
                .SetColumnStyle(3, xlNumStyle)
                RowCount = 3
                .SetCellValue(RowCount, 1, "Action")
                .SetCellValue(RowCount, 2, "Count")
                .SetCellValue(RowCount, 3, "Pct")
                .SetCellValue(RowCount, 4, "Actual Savings")
                .SetCellValue(RowCount, 5, "Potential Savings")
                .SetCellValue(RowCount, 6, "Average Delay (minutes)")
                .SetCellStyle(RowCount, 1, RowCount, 6, xlStyleHeader)
                For Each item As Reports.E28_VerificationItem In VerificationList.Values
                    RowCount += 1
                    .SetCellValue(RowCount, 1, item.Verification)
                    .SetCellValue(RowCount, 2, item.Count)
                    If VerificationList.Totals.Count <> 0 Then
                        .SetCellValue(RowCount, 3, item.Count / VerificationList.Totals.Count)
                    End If
                    .SetCellValue(RowCount, 4, item.ActualSavings)
                    .SetCellValue(RowCount, 5, item.PotentialSavings)
                    .SetCellValue(RowCount, 6, item.AverageDelay)
                Next
                RowCount += 1
                .SetCellValue(RowCount, 1, VerificationList.Totals.Verification)
                .SetCellValue(RowCount, 2, VerificationList.Totals.Count)
                .SetCellValue(RowCount, 4, VerificationList.Totals.ActualSavings)
                .SetCellValue(RowCount, 5, VerificationList.Totals.PotentialSavings)
                .SetCellValue(RowCount, 6, VerificationList.Totals.AverageDelay)
                .SetCellStyle(RowCount, 1, RowCount, 6, xlStyleLightGray)
                .AutoFitColumn(1, 6)
                RowCount = 1
                .SetCellValue(RowCount, 12, "Time Logged")
                .MergeWorksheetCells(RowCount, 12, RowCount, 35)
                RowCount = 2
                .SetCellValue(RowCount, 12, "Day")
                .MergeWorksheetCells(RowCount, 12, RowCount, 19)
                .SetCellValue(RowCount, 20, "Evening")
                .MergeWorksheetCells(RowCount, 20, RowCount, 24)
                .SetCellValue(RowCount, 25, "Night")
                .MergeWorksheetCells(RowCount, 25, RowCount, 32)
                .SetCellValue(RowCount, 33, "Morning")
                .MergeWorksheetCells(RowCount, 33, RowCount, 35)
                RowCount = 3
                .SetCellValue(RowCount, 8, "Reason")
                .SetCellValue(RowCount, 9, "Postponed")
                .SetCellValue(RowCount, 10, "Rejected")
                .SetCellValue(RowCount, 11, "Total")
                .SetCellStyle(RowCount, 8, RowCount, 11, xlStyleHeader)

                .SetCellValue(RowCount, 12, "09")
                .SetCellValue(RowCount, 13, "10")
                .SetCellValue(RowCount, 14, "11")
                .SetCellValue(RowCount, 15, "12")
                .SetCellValue(RowCount, 16, "13")
                .SetCellValue(RowCount, 17, "14")
                .SetCellValue(RowCount, 18, "15")
                .SetCellValue(RowCount, 19, "16")
                .SetCellValue(RowCount, 20, "17")
                .SetCellValue(RowCount, 21, "18")
                .SetCellValue(RowCount, 22, "19")
                .SetCellValue(RowCount, 23, "20")
                .SetCellValue(RowCount, 24, "21")
                .SetCellValue(RowCount, 25, "22")
                .SetCellValue(RowCount, 26, "23")
                .SetCellValue(RowCount, 27, "00")
                .SetCellValue(RowCount, 28, "01")
                .SetCellValue(RowCount, 29, "02")
                .SetCellValue(RowCount, 30, "03")
                .SetCellValue(RowCount, 31, "04")
                .SetCellValue(RowCount, 32, "05")
                .SetCellValue(RowCount, 33, "06")
                .SetCellValue(RowCount, 34, "07")
                .SetCellValue(RowCount, 35, "08")
                Dim xlTitle = .CreateStyle
                xlTitle.Font.Bold = True
                xlTitle.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center
                xlTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlTitle.Fill.SetPatternForegroundColor(Color.PeachPuff)
                .SetCellStyle(1, 12, 3, 35, xlTitle)
                Dim Others As New Reports.E28_ActionReasonItem("Other reasons")
                ActionReasonRow(ActionReasonList.Total)
                For Each item As Reports.E28_ActionReasonItem In ActionReasonList.GetSorted.Values
                    If item.Total > 1 Then
                        ActionReasonRow(item)
                    Else
                        Others.AddItem(item)
                    End If
                Next
                'RowCount += 1
                ActionReasonRow(Others)
                .DrawBorder(4, 8, 5, 35, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                .DrawBorder(1, 12, RowCount, 19, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                .DrawBorder(1, 20, RowCount, 24, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                .DrawBorder(1, 25, RowCount, 32, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                .DrawBorder(1, 33, RowCount, 35, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                .AutoFitColumn(8, 35)
            End With

            ThisDocument.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub ActionReasonRow(item As Reports.E28_ActionReasonItem)
        RowCount += 1
        ThisDocument.SetCellStyle(RowCount + 1, 11, RowCount + 1, 35, xlDecStyle)
        ThisDocument.SetCellValue(RowCount, 8, item.Reason)
        ThisDocument.SetCellValue(RowCount, 9, item.Postpone)
        ThisDocument.SetCellValue(RowCount, 10, item.Reject)
        ThisDocument.SetCellValue(RowCount, 11, item.Total)
        ThisDocument.SetCellValue(RowCount + 1, 8, "(average delay)")
        ThisDocument.SetCellValue(RowCount + 1, 11, item.AverageDelay)
        ThisDocument.SetCellStyle(RowCount + 1, 8, RowCount + 1, 11, xlStyleLightGray)
        For hourindex = 0 To 23
            If (item.PerHourItems(hourindex).Count > 0) Then

                ThisDocument.SetCellValue(RowCount, hourindex + 12, item.PerHourItems(hourindex).Count)
                ThisDocument.SetCellValue(RowCount + 1, hourindex + 12, item.PerHourItems(hourindex).DelayAverage)
                Dim xlCol = ThisDocument.CreateStyle
                xlCol.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlCol.Fill.SetPatternForegroundColor(Color.FromArgb(item.ValueColor(hourindex)))
                ThisDocument.SetCellStyle(RowCount + 1, hourindex + 12, xlCol)
            End If
        Next
        RowCount += 1
    End Sub
    Private Sub E28_AddVerificationWorksheet(wsName As String, agentlist As Reports.E28_GroupSubItemList, grouptotals As Boolean, withnames As Boolean)

        ThisDocument.AddWorksheet(wsName)
        Dim xlStyleHeader As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlStyleHeader.Font.Bold = True
        xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
        xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

        Dim xlStyleGroup As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlStyleGroup.Font.Bold = True
        xlStyleGroup.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleGroup.Fill.SetPatternForegroundColor(Color.LightGray)



        Dim xlNumStyle As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlNumStyle.FormatCode = "@"
        ThisDocument.SetColumnStyle(1, 19, xlNumStyle)
        xlNumStyle.FormatCode = "#,##0.00"
        ThisDocument.SetColumnStyle(5, 19, xlNumStyle)
        xlNumStyle.FormatCode = "#,##0"
        ThisDocument.SetColumnStyle(3, xlNumStyle)
        ThisDocument.SetColumnStyle(8, xlNumStyle)
        ThisDocument.SetColumnStyle(12, xlNumStyle)
        ThisDocument.SetColumnStyle(16, xlNumStyle)
        xlNumStyle.FormatCode = "#,##0.00%"
        ThisDocument.SetColumnStyle(4, xlNumStyle)
        ThisDocument.SetColumnStyle(9, xlNumStyle)
        ThisDocument.SetColumnStyle(13, xlNumStyle)
        RowCount = 1
        ThisDocument.FreezePanes(1, 2)
        ThisDocument.SetCellValue(RowCount, 1, "Group")
        ThisDocument.SetCellValue(RowCount, 2, "Agent")

        ThisDocument.SetCellValue(RowCount, 3, "Actioned")
        ThisDocument.SetCellValue(RowCount, 4, "Actioned PCT")
        ThisDocument.SetCellValue(RowCount, 5, "Actioned Actual Savings")
        ThisDocument.SetCellValue(RowCount, 6, "Actioned Potential Savings")
        ThisDocument.SetCellValue(RowCount, 7, "Actioned Average Delay")

        ThisDocument.SetCellValue(RowCount, 8, "Postponed")
        ThisDocument.SetCellValue(RowCount, 9, "Postponed PCT")
        ThisDocument.SetCellValue(RowCount, 10, "Postponed Potential Savings")
        ThisDocument.SetCellValue(RowCount, 11, "Postponed Average Delay")

        ThisDocument.SetCellValue(RowCount, 12, "Rejected")
        ThisDocument.SetCellValue(RowCount, 13, "Rejected PCT")
        ThisDocument.SetCellValue(RowCount, 14, "Rejected Potential Savings")
        ThisDocument.SetCellValue(RowCount, 15, "Rejected Average Delay")

        ThisDocument.SetCellValue(RowCount, 16, "Total")
        ThisDocument.SetCellValue(RowCount, 17, "Total Actual Savings")
        ThisDocument.SetCellValue(RowCount, 18, "Total Potential Savings")
        ThisDocument.SetCellValue(RowCount, 19, "Total Average Delay")
        ThisDocument.SetCellStyle(RowCount, 1, RowCount, 19, xlStyleHeader)
        For Each item As Reports.E28_GroupSubItemItem In agentlist.Values
            If (grouptotals And item.Name = "") Or (withnames And item.Name <> "") Then
                RowCount += 1
                ThisDocument.SetCellValue(RowCount, 1, item.Group)
                ThisDocument.SetCellValue(RowCount, 2, item.Name)

                ThisDocument.SetCellValue(RowCount, 3, item.Actioned.Count)
                If item.Total.Count <> 0 Then
                    ThisDocument.SetCellValue(RowCount, 4, item.Actioned.Count / item.Total.Count)
                End If
                ThisDocument.SetCellValue(RowCount, 5, item.Actioned.ActualSavings)
                ThisDocument.SetCellValue(RowCount, 6, item.Actioned.PotentialSaving)
                ThisDocument.SetCellValue(RowCount, 7, item.Actioned.MinutesDelayAverage)

                ThisDocument.SetCellValue(RowCount, 8, item.Postponed.Count)
                If item.Total.Count <> 0 Then
                    ThisDocument.SetCellValue(RowCount, 9, item.Postponed.Count / item.Total.Count)
                End If
                ThisDocument.SetCellValue(RowCount, 10, item.Postponed.PotentialSaving)
                ThisDocument.SetCellValue(RowCount, 11, item.Postponed.MinutesDelayAverage)

                ThisDocument.SetCellValue(RowCount, 12, item.Rejected.Count)
                If item.Total.Count <> 0 Then
                    ThisDocument.SetCellValue(RowCount, 13, item.Rejected.Count / item.Total.Count)
                End If
                ThisDocument.SetCellValue(RowCount, 14, item.Rejected.PotentialSaving)
                ThisDocument.SetCellValue(RowCount, 15, item.Rejected.MinutesDelayAverage)

                ThisDocument.SetCellValue(RowCount, 16, item.Total.Count)
                ThisDocument.SetCellValue(RowCount, 17, item.Total.ActualSavings)
                ThisDocument.SetCellValue(RowCount, 18, item.Total.PotentialSaving)
                ThisDocument.SetCellValue(RowCount, 19, item.Total.MinutesDelayAverage)
                If grouptotals And withnames And item.Name = "" Then
                    ThisDocument.SetCellStyle(RowCount, 1, RowCount, 19, xlStyleGroup)
                End If
            End If
        Next
        If Not (grouptotals And withnames) Then
            ThisDocument.Filter(1, 1, RowCount, 19)
            ThisDocument.Sort(2, 1, RowCount, 19, 16, False)
            ThisDocument.Sort(2, 1, RowCount, 19, 17, False)
        End If
        RowCount += 2
        ThisDocument.SetCellValue(RowCount, 1, "Grand Total")
        ThisDocument.SetCellValue(RowCount, 2, "")

        ThisDocument.SetCellValue(RowCount, 3, agentlist.Totals.Actioned.Count)
        If agentlist.Totals.Total.Count <> 0 Then
            ThisDocument.SetCellValue(RowCount, 4, agentlist.Totals.Actioned.Count / agentlist.Totals.Total.Count)
        End If
        ThisDocument.SetCellValue(RowCount, 5, agentlist.Totals.Actioned.ActualSavings)
        ThisDocument.SetCellValue(RowCount, 6, agentlist.Totals.Actioned.PotentialSaving)
        ThisDocument.SetCellValue(RowCount, 7, agentlist.Totals.Actioned.MinutesDelayAverage)

        ThisDocument.SetCellValue(RowCount, 8, agentlist.Totals.Postponed.Count)
        If agentlist.Totals.Total.Count <> 0 Then
            ThisDocument.SetCellValue(RowCount, 9, agentlist.Totals.Postponed.Count / agentlist.Totals.Total.Count)
        End If
        ThisDocument.SetCellValue(RowCount, 10, agentlist.Totals.Postponed.PotentialSaving)
        ThisDocument.SetCellValue(RowCount, 11, agentlist.Totals.Postponed.MinutesDelayAverage)

        ThisDocument.SetCellValue(RowCount, 12, agentlist.Totals.Rejected.Count)
        If agentlist.Totals.Total.Count <> 0 Then
            ThisDocument.SetCellValue(RowCount, 13, agentlist.Totals.Rejected.Count / agentlist.Totals.Total.Count)
        End If
        ThisDocument.SetCellValue(RowCount, 14, agentlist.Totals.Rejected.PotentialSaving)
        ThisDocument.SetCellValue(RowCount, 15, agentlist.Totals.Rejected.MinutesDelayAverage)

        ThisDocument.SetCellValue(RowCount, 16, agentlist.Totals.Total.Count)
        ThisDocument.SetCellValue(RowCount, 17, agentlist.Totals.Total.ActualSavings)
        ThisDocument.SetCellValue(RowCount, 18, agentlist.Totals.Total.PotentialSaving)
        ThisDocument.SetCellValue(RowCount, 19, agentlist.Totals.Total.MinutesDelayAverage)
        ThisDocument.SetCellStyle(RowCount, 1, RowCount, 19, xlStyleLightGray)

        ThisDocument.DrawBorder(1, 3, RowCount, 7, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
        ThisDocument.DrawBorder(1, 8, RowCount, 11, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
        ThisDocument.DrawBorder(1, 12, RowCount, 15, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
        ThisDocument.DrawBorder(1, 16, RowCount, 19, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
        ThisDocument.AutoFitColumn(1, 19)
    End Sub
    Private Sub E28_AddActionedOnlyWorksheet(wsName As String, agentlist As Reports.E28_GroupSubItemList)

        ThisDocument.AddWorksheet(wsName)
        Dim xlStyleHeader As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlStyleHeader.Font.Bold = True
        xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
        xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

        Dim xlStyleGroup As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlStyleGroup.Font.Bold = True
        xlStyleGroup.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
        xlStyleGroup.Fill.SetPatternForegroundColor(Color.LightGray)



        Dim xlNumStyle As SpreadsheetLight.SLStyle = ThisDocument.CreateStyle
        xlNumStyle.FormatCode = "@"
        ThisDocument.SetColumnStyle(1, 19, xlNumStyle)
        xlNumStyle.FormatCode = "#,##0.00"
        ThisDocument.SetColumnStyle(4, xlNumStyle)
        xlNumStyle.FormatCode = "#,##0"
        ThisDocument.SetColumnStyle(3, xlNumStyle)
        RowCount = 1
        ThisDocument.FreezePanes(1, 2)
        ThisDocument.SetCellValue(RowCount, 1, "Group")
        ThisDocument.SetCellValue(RowCount, 2, "Client")

        ThisDocument.SetCellValue(RowCount, 3, "Actioned")
        ThisDocument.SetCellValue(RowCount, 4, "Actioned Actual Savings")

        ThisDocument.SetCellStyle(RowCount, 1, RowCount, 4, xlStyleHeader)
        For Each item As Reports.E28_GroupSubItemItem In agentlist.Values
            RowCount += 1
            ThisDocument.SetCellValue(RowCount, 1, item.Group)
            ThisDocument.SetCellValue(RowCount, 2, item.Name)

            ThisDocument.SetCellValue(RowCount, 3, item.Actioned.Count)
            ThisDocument.SetCellValue(RowCount, 4, item.Actioned.ActualSavings)
            If item.Name = "" Then
                ThisDocument.SetCellStyle(RowCount, 1, RowCount, 4, xlStyleGroup)
            End If
        Next
        If Not True Then
            ThisDocument.Filter(1, 1, RowCount, 4)
            ThisDocument.Sort(2, 1, RowCount, 4, 4, False)
        End If
        RowCount += 2
        ThisDocument.SetCellValue(RowCount, 1, "Grand Total")
        ThisDocument.SetCellValue(RowCount, 2, "")

        ThisDocument.SetCellValue(RowCount, 3, agentlist.Totals.Actioned.Count)
        ThisDocument.SetCellValue(RowCount, 4, agentlist.Totals.Actioned.ActualSavings)
        ThisDocument.SetCellStyle(RowCount, 1, RowCount, 4, xlStyleLightGray)

        ThisDocument.DrawBorder(1, 3, RowCount, 4, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
        ThisDocument.AutoFitColumn(1, 4)
    End Sub
    Public Sub E30_AirTicketsFullDetails(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHC As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHC.Font.Italic = True
                xlStyleHC.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHC.Fill.SetPatternForegroundColor(Color.LightGray)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 67, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(25, 36, xlNumStyle)
                .SetColumnStyle(25, 26, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, xlNumStyle)
                .SetColumnStyle(15, xlNumStyle)
                .SetColumnStyle(47, 48, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(9, xlNumStyle)
                xlNumStyle.FormatCode = "0;-0;"
                .SetColumnStyle(40, xlNumStyle)

                .SetCellValue(1, 1, "Issue Date")
                .SetCellValue(1, 2, "Client Code")
                .SetCellValue(1, 3, "Client Name")
                .SetCellValue(1, 4, "Omit")
                .SetCellValue(1, 5, "Void")
                .SetCellValue(1, 6, "PNR")
                .SetCellValue(1, 7, "Ticket Number")
                .SetCellValue(1, 8, "Passenger")
                .SetCellValue(1, 9, "Pax Count")
                .SetCellValue(1, 10, "Product Type")
                .SetCellValue(1, 11, "Action Type")
                .SetCellValue(1, 12, "Inv Code")
                .SetCellValue(1, 13, "Inv Series")
                .SetCellValue(1, 14, "Inv Number")
                .SetCellValue(1, 15, "Invoice Date")
                .SetCellValue(1, 16, "Vessel")
                .SetCellValue(1, 17, "Booked By")
                .SetCellValue(1, 18, "Office/Dept")
                .SetCellValue(1, 19, "Reason For Travel")
                .SetCellValue(1, 20, "Cost Centre")

                .SetCellValue(1, 21, "Requisition Number")
                .SetCellValue(1, 22, "OPT")
                .SetCellValue(1, 23, "TRID/MarineFare")
                .SetCellValue(1, 24, "Account Code")

                .SetCellValue(1, 25, "FC")
                .SetCellValue(1, 26, "NetPayable FC")

                .SetCellValue(1, 27, "Net Payable")
                .SetCellValue(1, 28, "Net Purchase")
                .SetCellValue(1, 29, "Face Value")
                .SetCellValue(1, 30, "Taxes")
                .SetCellValue(1, 31, "Discount")
                .SetCellValue(1, 32, "Commission")
                .SetCellValue(1, 33, "Cancellation Fee")

                .SetCellValue(1, 34, "FV Extra")
                .SetCellValue(1, 35, "Tax Extra")
                .SetCellValue(1, 36, "Service Fee")

                .SetCellValue(1, 37, "Verified")
                .SetCellValue(1, 38, "Remarks")
                .SetCellValue(1, 39, "Transaction Type")
                .SetCellValue(1, 40, "RegNr")
                .SetCellValue(1, 41, "Ticketing Airline")
                .SetCellValue(1, 42, "Routing")
                .SetCellValue(1, 43, "SalesPerson")
                .SetCellValue(1, 44, "Issuing Agent")
                .SetCellValue(1, 45, "Creator Agent")
                .SetCellValue(1, 46, "Reference")
                .SetCellValue(1, 47, "Departure Date")
                .SetCellValue(1, 48, "Arrival Date")
                .SetCellValue(1, 49, "Connected Document")
                .SetCellValue(1, 50, "Pax Remarks")
                .SetCellValue(1, 51, "Rank")
                .SetCellValue(1, 52, "Nationality")
                .SetCellValue(1, 53, "REF1")
                .SetCellValue(1, 54, "REF2")
                .SetCellValue(1, 55, "REF3")
                .SetCellValue(1, 56, "REF4")
                .SetCellValue(1, 57, "REF5")
                .SetCellValue(1, 58, "REF6")
                .SetCellValue(1, 59, "REF7")
                .SetCellValue(1, 60, "REF8")
                .SetCellValue(1, 61, "REF9")
                .SetCellValue(1, 62, "REF10")
                .SetCellValue(1, 63, "REF11")
                .SetCellValue(1, 64, "REF12")
                .SetCellValue(1, 65, "REF13")
                .SetCellValue(1, 66, "REF14")
                .SetCellValue(1, 67, "REF15")
                .SetCellValue(1, 68, "Net Remit")

                .SetCellStyle(1, 1, 1, 68, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To 10
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(13) <> 0 Then
                        For j = 11 To 14
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                    End If
                    For j = 15 To 23
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    .SetCellValue(xlINVCount, 25, mdsDataSet.Tables(0).Rows(i).Item(66))
                    .SetCellValue(xlINVCount, 26, mdsDataSet.Tables(0).Rows(i).Item(65))
                    If mdsDataSet.Tables(0).Rows(i).Item(67) = 1 Then
                        .SetCellStyle(xlINVCount, 25, xlINVCount, 26, xlStyleHC)
                    End If
                    For j = 24 To 64
                        .SetCellValue(xlINVCount, j + 3, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    .SetCellValue(xlINVCount, 68, mdsDataSet.Tables(0).Rows(i).Item(68))
                    If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 68, xlStyleOmit)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 68, xlStyleVoid)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 68, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 68)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E36_SeaChefs_AllUnits(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0
        Dim xlHeaderID As Integer = 0
        Dim pInvNumber As String = ""
        Dim pInvRow As Integer = 0
        Dim pInvSum As Decimal = 0

        Try
            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Create Invoices")

                .FreezePanes(1, 0)

                Dim xlStyleHeaderNotes As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeaderNotes.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeaderNotes.Fill.SetPatternForegroundColor(Color.FromArgb(255, 169, 208, 142))
                xlStyleHeaderNotes.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleFixed As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleFixed.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleFixed.Fill.SetPatternForegroundColor(Color.FromArgb(255, 142, 169, 219))
                xlStyleFixed.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 0, 204, 255))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleGrey As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGrey.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGrey.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleGrey.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTitle.Font.Bold = True
                xlStyleTitle.Font.FontSize = 15
                xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleTitle.Fill.SetPatternForegroundColor(Color.FromArgb(255, 191, 191, 191))
                xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 32, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(9, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(10, 12, xlNumStyle)

                .SetCellValue(1, 1, "Business Unit")
                .SetCellValue(1, 2, "Vessel")
                .SetCellValue(1, 3, "ClientCode")
                .SetCellValue(1, 4, "ClientName")
                .SetCellValue(1, 5, "InvCode")
                .SetCellValue(1, 6, "InvSeries")
                .SetCellValue(1, 7, "Invoice Number")
                .SetCellValue(1, 8, "Invoice Currency")
                .SetCellValue(1, 9, "Invoice Amount")
                .SetCellValue(1, 10, "Invoice Date")
                .SetCellValue(1, 11, "FromDate")
                .SetCellValue(1, 12, "ToDate")
                .SetCellValue(1, 13, "Line")
                .SetCellValue(1, 14, "Description")
                .SetCellValue(1, 15, "Distribution Combination")
                .SetCellValue(1, 16, "PNR")
                .SetCellValue(1, 17, "TicketNumber")
                .SetCellValue(1, 18, "PaxCount")
                .SetCellValue(1, 19, "ProductType")
                .SetCellValue(1, 20, "ActionType")
                .SetCellValue(1, 21, "BookedBy")
                .SetCellValue(1, 22, "Office")
                .SetCellValue(1, 23, "ReasonForTravel")
                .SetCellValue(1, 24, "RequisitionNumber")
                .SetCellValue(1, 25, "OPT")
                .SetCellValue(1, 26, "TRID-MarineFare")
                .SetCellValue(1, 27, "Verified")
                .SetCellValue(1, 28, "Remarks")
                .SetCellValue(1, 29, "RegNr")
                .SetCellValue(1, 30, "SalesPerson")
                .SetCellValue(1, 31, "IssuingAgent")
                .SetCellValue(1, 32, "CreatorAgent")


                .SetCellStyle(1, 1, 1, 32, xlStyleHeader)

                xlINVCount = 1
                pInvNumber = mdsDataSet.Tables(0).Rows(0).Item(6)
                pInvRow = xlINVCount + 1
                pInvSum = 0
                xlHeaderID = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    If mdsDataSet.Tables(0).Rows(i).Item(6) = pInvNumber Then
                        pInvSum += mdsDataSet.Tables(0).Rows(i).Item(8)
                    Else
                        xlHeaderID += 1
                        .SetCellValue(pInvRow, 9, pInvSum)
                        For k = pInvRow + 1 To xlINVCount - 1
                            .SetCellStyle(k, 5, k, 9, xlStyleGrey)

                            .SetCellValue(k, 5, "")
                            .SetCellValue(k, 5, "")
                            .SetCellValue(k, 7, "")
                            .SetCellValue(k, 8, "")
                            .SetCellValue(k, 9, "")
                        Next
                        pInvSum = mdsDataSet.Tables(0).Rows(i).Item(8)
                        pInvNumber = mdsDataSet.Tables(0).Rows(i).Item(6)
                        pInvRow = xlINVCount
                    End If
                    .SetCellValue(xlINVCount, 5, xlHeaderID)
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item(2))
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item(6))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item(14))
                    .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item(15))
                    .SetCellValue(xlINVCount, 17, mdsDataSet.Tables(0).Rows(i).Item(16))
                    .SetCellValue(xlINVCount, 18, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(xlINVCount, 19, mdsDataSet.Tables(0).Rows(i).Item(18))
                    .SetCellValue(xlINVCount, 20, mdsDataSet.Tables(0).Rows(i).Item(19))
                    .SetCellValue(xlINVCount, 21, mdsDataSet.Tables(0).Rows(i).Item(20))
                    .SetCellValue(xlINVCount, 22, mdsDataSet.Tables(0).Rows(i).Item(21))
                    .SetCellValue(xlINVCount, 23, mdsDataSet.Tables(0).Rows(i).Item(22))
                    .SetCellValue(xlINVCount, 24, mdsDataSet.Tables(0).Rows(i).Item(23))
                    .SetCellValue(xlINVCount, 25, mdsDataSet.Tables(0).Rows(i).Item(24))
                    .SetCellValue(xlINVCount, 26, mdsDataSet.Tables(0).Rows(i).Item(25))
                    .SetCellValue(xlINVCount, 27, mdsDataSet.Tables(0).Rows(i).Item(26))
                    .SetCellValue(xlINVCount, 28, mdsDataSet.Tables(0).Rows(i).Item(27))
                    .SetCellValue(xlINVCount, 29, mdsDataSet.Tables(0).Rows(i).Item(28))
                    .SetCellValue(xlINVCount, 30, mdsDataSet.Tables(0).Rows(i).Item(29))
                    .SetCellValue(xlINVCount, 31, mdsDataSet.Tables(0).Rows(i).Item(30))
                    .SetCellValue(xlINVCount, 32, mdsDataSet.Tables(0).Rows(i).Item(31))


                Next
                If pInvNumber <> "" AndAlso pInvSum <> 0 Then
                    .SetCellValue(pInvRow, 9, pInvSum)
                    For k = pInvRow + 1 To xlINVCount
                        .SetCellStyle(k, 5, k, 9, xlStyleGrey)

                        .SetCellValue(k, 5, "")
                        .SetCellValue(k, 5, "")
                        .SetCellValue(k, 7, "")
                        .SetCellValue(k, 8, "")
                        .SetCellValue(k, 9, "")
                    Next
                End If
                .AutoFitColumn(1, 32)


            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E38_AirTicketSales(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 42, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(9, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, 2, xlNumStyle)
                .SetColumnStyle(11, xlNumStyle)
                .SetColumnStyle(35, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(18, xlNumStyle)
                xlNumStyle.FormatCode = "0;-0;"
                .SetColumnStyle(30, xlNumStyle)

                .SetCellValue(1, 1, "Invoice Date")
                .SetCellValue(1, 2, "Departure Date")
                .SetCellValue(1, 3, "Booked By")
                .SetCellValue(1, 4, "Invoice Code")
                .SetCellValue(1, 5, "Vessel")
                .SetCellValue(1, 6, "Passenger")
                .SetCellValue(1, 7, "Routing")
                .SetCellValue(1, 8, "Ticketing Airline")
                .SetCellValue(1, 9, "Net Payable")
                .SetCellValue(1, 10, "Reason For Travel")
                .SetCellValue(1, 11, "Issue Date")
                .SetCellValue(1, 12, "Client Code")
                .SetCellValue(1, 13, "Client Name")
                .SetCellValue(1, 14, "Omit")
                .SetCellValue(1, 15, "Void")
                .SetCellValue(1, 16, "PNR")
                .SetCellValue(1, 17, "Ticket Number")
                .SetCellValue(1, 18, "Pax Count")
                .SetCellValue(1, 19, "Product Type")
                .SetCellValue(1, 20, "Action Type")
                .SetCellValue(1, 21, "Office/Dept")
                .SetCellValue(1, 22, "Cost Centre")

                .SetCellValue(1, 23, "Requisition Number")
                .SetCellValue(1, 24, "OPT")
                .SetCellValue(1, 25, "TRID/MarineFare")
                .SetCellValue(1, 26, "Account Code")

                .SetCellValue(1, 27, "Verified")
                .SetCellValue(1, 28, "Remarks")
                .SetCellValue(1, 29, "Transaction Type")
                .SetCellValue(1, 30, "RegNr")
                .SetCellValue(1, 31, "SalesPerson")
                .SetCellValue(1, 32, "Issuing Agent")
                .SetCellValue(1, 33, "Creator Agent")
                .SetCellValue(1, 34, "Reference")
                .SetCellValue(1, 35, "Arrival Date")
                .SetCellValue(1, 36, "Connected Document")
                .SetCellValue(1, 37, "Pax Remarks")
                .SetCellValue(1, 38, "Invoice Status")
                .SetCellValue(1, 39, "Other Services")
                .SetCellValue(1, 40, "Client Team")
                .SetCellValue(1, 41, "Supplier Code")
                .SetCellValue(1, 42, "Supplier Name")

                .SetCellStyle(1, 1, 1, 42, xlStyleHeader)

                xlINVCount = 1
                Dim inv As String
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1

                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(14)) '.SetCellValue(1, 36, "Invoice Date")
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item(35)) '.SetCellValue(1, 36, "Departure Date")
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item(16)) '.SetCellValue(1, 17, "Booked By")
                    If mdsDataSet.Tables(0).Rows(i).Item(13) <> 0 Then
                        inv = $"{mdsDataSet.Tables(0).Rows(i).Item(11)} {mdsDataSet.Tables(0).Rows(i).Item(12)} {mdsDataSet.Tables(0).Rows(i).Item(13)}".Replace("  ", " ").Trim
                    Else
                        inv = ""
                    End If
                    .SetCellValue(xlINVCount, 4, inv) '.SetCellValue(1, 12, "Inv Code") &.SetCellValue(1, 13, "Inv Series") & .SetCellValue(1, 14, "Inv Number")
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(15)) '.SetCellValue(1, 16, "Vessel")
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(7)) '.SetCellValue(1, 8, "Passenger")
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item(30)) '.SetCellValue(1, 31, "Routing")
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(29)) '.SetCellValue(1, 30, "Ticketing Airline")
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(24)) '.SetCellValue(1, 25, "Net Payable")
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item(18)) '.SetCellValue(1, 19, "Reason For Travel")

                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(0))
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item(1))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(2))
                    .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(xlINVCount, 17, mdsDataSet.Tables(0).Rows(i).Item(6))
                    .SetCellValue(xlINVCount, 18, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(xlINVCount, 19, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(xlINVCount, 20, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(xlINVCount, 21, mdsDataSet.Tables(0).Rows(i).Item(17))

                    For j = 19 To 23
                        .SetCellValue(xlINVCount, j + 3, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    For j = 25 To 28
                        .SetCellValue(xlINVCount, j + 2, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    For j = 31 To 34
                        .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    For j = 35 To 37
                        .SetCellValue(xlINVCount, j, mdsDataSet.Tables(0).Rows(i).Item(j + 1))
                    Next

                    If CInt(mdsDataSet.Tables(0).Rows(i).Item(39)) = 43 Then
                        .SetCellValue(xlINVCount, 38, "Cancelled")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleCancelled)
                    ElseIf CStr(mdsDataSet.Tables(0).Rows(i).Item(40)) <> "" Then
                        .SetCellValue(xlINVCount, 38, $"Cancels {CStr(mdsDataSet.Tables(0).Rows(i).Item(40))}")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleCancelled)
                    End If
                    .SetCellValue(xlINVCount, 39, mdsDataSet.Tables(0).Rows(i).Item(41))
                    .SetCellValue(xlINVCount, 40, mdsDataSet.Tables(0).Rows(i).Item(42))
                    .SetCellValue(xlINVCount, 41, mdsDataSet.Tables(0).Rows(i).Item("SuppCode"))
                    .SetCellValue(xlINVCount, 42, mdsDataSet.Tables(0).Rows(i).Item("SuppName"))
                    If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleOmit)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleVoid)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 38, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 42)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E42_AirTicketsWithFC(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHC As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHC.Font.Italic = True
                xlStyleHC.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHC.Fill.SetPatternForegroundColor(Color.LightGray)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 10, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(8, 9, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(3, xlNumStyle)

                .SetCellValue(1, 1, "Client Code")
                .SetCellValue(1, 2, "Client Name")
                .SetCellValue(1, 3, "Issue Date")
                .SetCellValue(1, 4, "Product Type")
                .SetCellValue(1, 5, "Action Type")
                .SetCellValue(1, 6, "Inv Code")
                .SetCellValue(1, 7, "Inv Number")

                .SetCellValue(1, 8, "Net Payable")
                .SetCellValue(1, 9, "NetPayable FC")
                .SetCellValue(1, 10, "FC")

                .SetCellStyle(1, 1, 1, 10, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To 9
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(4) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 10, xlStyleRefund)
                    End If

                    If mdsDataSet.Tables(0).Rows(i).Item(10) = 1 Then
                        .SetCellStyle(xlINVCount, 9, xlINVCount, 10, xlStyleHC)
                    End If
                Next

                .AutoFitColumn(1, 11)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E19aProfitReportInvoicesTotal(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Profit Totals")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleYellow As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleYellow.Font.Italic = True
                xlStyleYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleYellow.Fill.SetPatternForegroundColor(Color.Yellow)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 14, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(2, 14, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(5, 5, xlNumStyle)
                .SetColumnStyle(9, 9, xlNumStyle)
                .SetColumnStyle(13, 13, xlNumStyle)

                .SetCellValue(1, 1, "Client Group")
                .SetCellValue(1, 2, "Net Payable Invoices")
                .SetCellValue(1, 3, "Net Buy Invoices")
                .SetCellValue(1, 4, "Profit Invoices")
                .SetCellValue(1, 5, "Pax Invoices")
                .SetCellValue(1, 6, "Net Payable CN")
                .SetCellValue(1, 7, "Net Buy CN")
                .SetCellValue(1, 8, "Profit CN")
                .SetCellValue(1, 9, "Pax CN")
                .SetCellValue(1, 10, "Net Payable")
                .SetCellValue(1, 11, "Net Buy")
                .SetCellValue(1, 12, "Profit")
                .SetCellValue(1, 13, "Pax")
                .SetCellValue(1, 14, "% CN Pax/Total Pax")

                .SetCellStyle(1, 1, 1, 14, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item(2))
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(1) - mdsDataSet.Tables(0).Rows(i).Item(2))
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(4) - mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(6))
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item(1) + mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(2) + mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(xlINVCount, 12, (mdsDataSet.Tables(0).Rows(i).Item(1) - mdsDataSet.Tables(0).Rows(i).Item(2)) + (mdsDataSet.Tables(0).Rows(i).Item(4) - mdsDataSet.Tables(0).Rows(i).Item(5)))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(3) + mdsDataSet.Tables(0).Rows(i).Item(6))
                    If mdsDataSet.Tables(0).Rows(i).Item(3) <> 0 Then
                        .SetCellValue(xlINVCount, 14, -mdsDataSet.Tables(0).Rows(i).Item(6) / mdsDataSet.Tables(0).Rows(i).Item(3) * 100)
                    End If

                Next
                .Sort(2, 1, xlINVCount, 14, 12, False)
                .AutoFitColumn(1, 14)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E49_Optimization_Monthly_Report(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Per Agent")
                .FreezePanes(3, 0)
                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleYellow As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleYellow.Fill.SetPatternForegroundColor(Color.Yellow)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlTextStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlTextStyle.FormatCode = "@"
                .SetColumnStyle(1, xlTextStyle)
                Dim xlDecimalStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlDecimalStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(5, xlDecimalStyle)
                Dim xlIntStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlIntStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(2, 4, xlIntStyle)
                Dim xlPctStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlPctStyle.FormatCode = "#,##0.00%;-#,##0.00%;"
                .SetColumnStyle(6, 8, xlPctStyle)
                Dim xlSavingsStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlSavingsStyle.FormatCode = "#,##0.00;-#,##0.00;"
                xlSavingsStyle.Font.Bold = True
                xlSavingsStyle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlSavingsStyle.Fill.SetPatternForegroundColor(Color.Aqua)
                xlSavingsStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right)

                .SetCellValue(1, 1, "Total")
                .SetCellValue(2, 6, "Percentage")
                .MergeWorksheetCells(2, 6, 2, 8)
                .SetCellValue(3, 1, "Agent")
                .SetCellValue(3, 2, "Rejected")
                .SetCellValue(3, 3, "Postponed")
                .SetCellValue(3, 4, "Actioned")
                .SetCellValue(3, 5, "Savings Amount")
                .SetCellValue(3, 6, "Rejected")
                .SetCellValue(3, 7, "Postponed")
                .SetCellValue(3, 8, "Actioned")

                .SetCellStyle(1, 1, 3, 8, xlStyleHeader)

                .AddWorksheet("Per Client")
                .FreezePanes(3, 0)
                .SetColumnStyle(1, 2, xlTextStyle)
                .SetColumnStyle(6, xlDecimalStyle)
                .SetColumnStyle(3, 5, xlIntStyle)
                .SetColumnStyle(7, 9, xlPctStyle)

                .SetCellValue(1, 1, "Total")
                .SetCellValue(2, 7, "Percentage")
                .MergeWorksheetCells(2, 7, 2, 9)
                .SetCellValue(3, 1, "Client Code")
                .SetCellValue(3, 2, "Client Name")
                .SetCellValue(3, 3, "Rejected")
                .SetCellValue(3, 4, "Postponed")
                .SetCellValue(3, 5, "Actioned")
                .SetCellValue(3, 6, "Savings Amount")
                .SetCellValue(3, 7, "Rejected")
                .SetCellValue(3, 8, "Postponed")
                .SetCellValue(3, 9, "Actioned")

                .SetCellStyle(1, 1, 3, 9, xlStyleHeader)

                Dim AgentCount As Integer = 3
                Dim ClientCount As Integer = 3
                Dim AgentSum(3) As Decimal
                Dim ClientSum(3) As Decimal
                Dim IsAgent As Boolean = False
                Dim RowTot As Decimal
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "Agent" Then
                        If Not IsAgent Then
                            .SelectWorksheet("Per Agent")
                            IsAgent = True
                        End If
                        AgentCount += 1
                        .SetCellValue(AgentCount, 1, mdsDataSet.Tables(0).Rows(i).Item(1))
                        .SetCellValue(AgentCount, 2, mdsDataSet.Tables(0).Rows(i).Item(3))
                        .SetCellValue(AgentCount, 3, mdsDataSet.Tables(0).Rows(i).Item(4))
                        .SetCellValue(AgentCount, 4, mdsDataSet.Tables(0).Rows(i).Item(5))
                        .SetCellValue(AgentCount, 5, mdsDataSet.Tables(0).Rows(i).Item(6))
                        RowTot = mdsDataSet.Tables(0).Rows(i).Item(3) + mdsDataSet.Tables(0).Rows(i).Item(4) + mdsDataSet.Tables(0).Rows(i).Item(5)
                        If RowTot <> 0 Then
                            .SetCellValue(AgentCount, 6, mdsDataSet.Tables(0).Rows(i).Item(3) / RowTot)
                            .SetCellValue(AgentCount, 7, mdsDataSet.Tables(0).Rows(i).Item(4) / RowTot)
                            .SetCellValue(AgentCount, 8, mdsDataSet.Tables(0).Rows(i).Item(5) / RowTot)
                            If mdsDataSet.Tables(0).Rows(i).Item(3) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(3) >= mdsDataSet.Tables(0).Rows(i).Item(4) And mdsDataSet.Tables(0).Rows(i).Item(3) >= mdsDataSet.Tables(0).Rows(i).Item(5) Then
                                .SetCellStyle(AgentCount, 6, xlStyleYellow)
                            End If
                            If mdsDataSet.Tables(0).Rows(i).Item(4) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(4) >= mdsDataSet.Tables(0).Rows(i).Item(3) And mdsDataSet.Tables(0).Rows(i).Item(4) >= mdsDataSet.Tables(0).Rows(i).Item(5) Then
                                .SetCellStyle(AgentCount, 7, xlStyleYellow)
                            End If
                            If mdsDataSet.Tables(0).Rows(i).Item(5) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(5) >= mdsDataSet.Tables(0).Rows(i).Item(3) And mdsDataSet.Tables(0).Rows(i).Item(5) >= mdsDataSet.Tables(0).Rows(i).Item(4) Then
                                .SetCellStyle(AgentCount, 8, xlStyleYellow)
                            End If
                        End If
                        AgentSum(0) += mdsDataSet.Tables(0).Rows(i).Item(3)
                        AgentSum(1) += mdsDataSet.Tables(0).Rows(i).Item(4)
                        AgentSum(2) += mdsDataSet.Tables(0).Rows(i).Item(5)
                        AgentSum(3) += mdsDataSet.Tables(0).Rows(i).Item(6)
                    Else
                        If IsAgent Then
                            .SelectWorksheet("Per Client")
                            IsAgent = False
                        End If
                        ClientCount += 1
                        .SetCellValue(ClientCount, 1, mdsDataSet.Tables(0).Rows(i).Item(1))
                        .SetCellValue(ClientCount, 2, mdsDataSet.Tables(0).Rows(i).Item(2))
                        .SetCellValue(ClientCount, 3, mdsDataSet.Tables(0).Rows(i).Item(3))
                        .SetCellValue(ClientCount, 4, mdsDataSet.Tables(0).Rows(i).Item(4))
                        .SetCellValue(ClientCount, 5, mdsDataSet.Tables(0).Rows(i).Item(5))
                        .SetCellValue(ClientCount, 6, mdsDataSet.Tables(0).Rows(i).Item(6))
                        RowTot = mdsDataSet.Tables(0).Rows(i).Item(3) + mdsDataSet.Tables(0).Rows(i).Item(4) + mdsDataSet.Tables(0).Rows(i).Item(5)
                        If RowTot <> 0 Then
                            .SetCellValue(ClientCount, 7, mdsDataSet.Tables(0).Rows(i).Item(3) / RowTot)
                            .SetCellValue(ClientCount, 8, mdsDataSet.Tables(0).Rows(i).Item(4) / RowTot)
                            .SetCellValue(ClientCount, 9, mdsDataSet.Tables(0).Rows(i).Item(5) / RowTot)
                            If mdsDataSet.Tables(0).Rows(i).Item(3) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(3) >= mdsDataSet.Tables(0).Rows(i).Item(4) And mdsDataSet.Tables(0).Rows(i).Item(3) >= mdsDataSet.Tables(0).Rows(i).Item(5) Then
                                .SetCellStyle(ClientCount, 7, xlStyleYellow)
                            End If
                            If mdsDataSet.Tables(0).Rows(i).Item(4) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(4) >= mdsDataSet.Tables(0).Rows(i).Item(3) And mdsDataSet.Tables(0).Rows(i).Item(4) >= mdsDataSet.Tables(0).Rows(i).Item(5) Then
                                .SetCellStyle(ClientCount, 8, xlStyleYellow)
                            End If
                            If mdsDataSet.Tables(0).Rows(i).Item(5) <> 0 And mdsDataSet.Tables(0).Rows(i).Item(5) >= mdsDataSet.Tables(0).Rows(i).Item(3) And mdsDataSet.Tables(0).Rows(i).Item(5) >= mdsDataSet.Tables(0).Rows(i).Item(4) Then
                                .SetCellStyle(ClientCount, 9, xlStyleYellow)
                            End If

                        End If
                        ClientSum(0) += mdsDataSet.Tables(0).Rows(i).Item(3)
                        ClientSum(1) += mdsDataSet.Tables(0).Rows(i).Item(4)
                        ClientSum(2) += mdsDataSet.Tables(0).Rows(i).Item(5)
                        ClientSum(3) += mdsDataSet.Tables(0).Rows(i).Item(6)
                    End If
                Next
                .SelectWorksheet("Per Agent")
                .SetCellValue(1, 2, AgentSum(0))
                .SetCellValue(1, 3, AgentSum(1))
                .SetCellValue(1, 4, AgentSum(2))
                .SetCellValue(1, 5, AgentSum(3))
                RowTot = AgentSum(0) + AgentSum(1) + AgentSum(2)
                If RowTot <> 0 Then
                    .SetCellValue(1, 6, AgentSum(0) / RowTot)
                    .SetCellValue(1, 7, AgentSum(1) / RowTot)
                    .SetCellValue(1, 8, AgentSum(2) / RowTot)
                    If AgentSum(0) <> 0 And AgentSum(0) >= AgentSum(1) And AgentSum(0) >= AgentSum(2) Then
                        .SetCellStyle(1, 6, xlStyleYellow)
                    End If
                    If AgentSum(1) <> 0 And AgentSum(1) >= AgentSum(0) And AgentSum(1) >= AgentSum(2) Then
                        .SetCellStyle(1, 7, xlStyleYellow)
                    End If
                    If AgentSum(2) <> 0 And AgentSum(2) >= AgentSum(0) And AgentSum(2) >= AgentSum(1) Then
                        .SetCellStyle(1, 8, xlStyleYellow)
                    End If

                End If
                .SetCellStyle(4, 5, AgentCount, 5, xlSavingsStyle)
                .AutoFitColumn(1, 8)
                .SelectWorksheet("Per Client")
                .SetCellValue(1, 3, ClientSum(0))
                .SetCellValue(1, 4, ClientSum(1))
                .SetCellValue(1, 5, ClientSum(2))
                .SetCellValue(1, 6, ClientSum(3))
                RowTot = ClientSum(0) + ClientSum(1) + ClientSum(2)
                If RowTot <> 0 Then
                    .SetCellValue(1, 7, ClientSum(0) / RowTot)
                    .SetCellValue(1, 8, ClientSum(1) / RowTot)
                    .SetCellValue(1, 9, ClientSum(2) / RowTot)
                    If ClientSum(0) <> 0 And ClientSum(0) >= ClientSum(1) And ClientSum(0) >= ClientSum(2) Then
                        .SetCellStyle(1, 7, xlStyleYellow)
                    End If
                    If ClientSum(1) <> 0 And ClientSum(1) >= ClientSum(0) And ClientSum(1) >= ClientSum(2) Then
                        .SetCellStyle(1, 8, xlStyleYellow)
                    End If
                    If ClientSum(2) <> 0 And ClientSum(2) >= ClientSum(0) And ClientSum(2) >= ClientSum(1) Then
                        .SetCellStyle(1, 9, xlStyleYellow)
                    End If
                End If
                .SetCellStyle(4, 6, ClientCount, 6, xlSavingsStyle)
                .AutoFitColumn(1, 9)
            End With


            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E50_Optimization_Annual_Report_by_Month(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Per Agent")
                .AddWorksheet("Per Client")

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleLightGreen As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleLightGreen.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleLightGreen.Fill.SetPatternForegroundColor(Color.LightGreen)

                Dim xlStyleLightSalmon As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleLightSalmon.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleLightSalmon.Fill.SetPatternForegroundColor(Color.LightSalmon)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlTextStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlTextStyle.FormatCode = "@"
                Dim xlDecimalStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlDecimalStyle.FormatCode = "#,##0.00;-#,##0.00;"
                Dim xlIntStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlIntStyle.FormatCode = "#,##0;-#,##0;"

                For iw = 0 To 1
                    If iw = 0 Then
                        .SelectWorksheet("Per Agent")
                    Else
                        .SelectWorksheet("Per Client")
                    End If
                    .FreezePanes(3, 0)
                    .SetColumnStyle(1, 2, xlTextStyle)
                    .SetColumnStyle(3, 54, xlIntStyle)
                    .SetColumnStyle(6, xlDecimalStyle)

                    .SetCellValue(1, 3, mReport.BSPMonthDate)
                    .MergeWorksheetCells(1, 3, 1, 6)
                    .SetCellValue(2, 2, "Total")
                    .SetCellValue(3, 3, "Rejected")
                    .SetCellValue(3, 4, "Postponed")
                    .SetCellValue(3, 5, "Actioned")
                    .SetCellValue(3, 6, "Savings Amount")
                    For j = 7 To 54 Step 4
                        .SetCellValue(1, j, MonthNames((j - 3) / 4 - 1))
                        .MergeWorksheetCells(1, j, 1, j + 3)
                        .SetCellValue(3, j, $"Rejected {(j - 3) / 4}")
                        .SetCellValue(3, j + 1, $"Postponed {(j - 3) / 4 }")
                        .SetCellValue(3, j + 2, $"Actioned {(j - 3) / 4 }")
                        .SetCellValue(3, j + 3, $"Savings Amount {(j - 3) / 4 }")
                        .SetColumnStyle(j + 3, xlDecimalStyle)
                    Next
                    .SetCellStyle(1, 1, 3, 54, xlStyleHeader)
                Next

                Dim AgentCount As Integer = 3
                Dim ClientCount As Integer = 3
                Dim AgentSum(54) As Decimal
                Dim ClientSum(54) As Decimal
                Dim IsAgent As Boolean = False
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If mdsDataSet.Tables(0).Rows(i).Item(0) = "Agent" Then
                        If Not IsAgent Then
                            .SelectWorksheet("Per Agent")
                            IsAgent = True
                        End If
                        AgentCount += 1
                        .SetCellValue(AgentCount, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                        For j = 3 To 54
                            .SetCellValue(AgentCount, j, mdsDataSet.Tables(0).Rows(i).Item(j))
                            AgentSum(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        Next
                    Else
                        If IsAgent Then
                            .SelectWorksheet("Per Client")
                            IsAgent = False
                        End If
                        ClientCount += 1
                        .SetCellValue(ClientCount, 1, mdsDataSet.Tables(0).Rows(i).Item(1))
                        .SetCellValue(ClientCount, 2, mdsDataSet.Tables(0).Rows(i).Item(2))
                        For j = 3 To 54
                            .SetCellValue(ClientCount, j, mdsDataSet.Tables(0).Rows(i).Item(j))
                            ClientSum(j) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        Next
                    End If
                Next
                .SelectWorksheet("Per Agent")
                For j = 3 To 54 Step 4
                    .SetCellValue(2, j, AgentSum(j))
                    .SetCellValue(2, j + 1, AgentSum(j + 1))
                    .SetCellValue(2, j + 2, AgentSum(j + 2))
                    .SetCellValue(2, j + 3, AgentSum(j + 3))
                    .SetCellStyle(4, j, AgentCount, j + 2, xlStyleLightSalmon)
                    .SetCellStyle(4, j + 3, AgentCount, j + 3, xlStyleLightGreen)
                    .DrawBorder(1, j, AgentCount, j + 3, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)
                Next
                .AutoFitColumn(1, 54)
                .SelectWorksheet("Per Client")
                For j = 3 To 54 Step 4
                    .SetCellValue(2, j, ClientSum(j))
                    .SetCellValue(2, j + 1, ClientSum(j + 1))
                    .SetCellValue(2, j + 2, ClientSum(j + 2))
                    .SetCellValue(2, j + 3, ClientSum(j + 3))
                    .SetCellStyle(4, j, ClientCount, j + 2, xlStyleLightSalmon)
                    .SetCellStyle(4, j + 3, ClientCount, j + 3, xlStyleLightGreen)
                    .DrawBorder(1, j, ClientCount, j + 3, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick)

                Next
                .AutoFitColumn(1, 54)
            End With


            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub E51_Daily_Profit_Totals_per_Category(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0


        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Daily Profit Report")
                .FreezePanes(5, 4)
                Dim xlStyleTotalCol As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTotalCol.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleTotalCol.Fill.SetPatternForegroundColor(Color.FromArgb(217, 225, 242)) ' Blue Accent 1, Lighter 80%

                Dim xlStyleAir1 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleAir1.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleAir1.Fill.SetPatternForegroundColor(Color.FromArgb(226, 239, 218)) ' Green Accent 6, Lighter 80%

                Dim xlStyleAir2 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleAir2.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleAir2.Fill.SetPatternForegroundColor(Color.FromArgb(198, 224, 180)) ' Green Accent 6, Lighter 60%

                Dim xlStyleServices1 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleServices1.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleServices1.Fill.SetPatternForegroundColor(Color.FromArgb(255, 242, 204)) ' Gold Accent 4, Lighter 80%

                Dim xlStyleServices2 As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleServices2.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleServices2.Fill.SetPatternForegroundColor(Color.FromArgb(255, 230, 153)) ' Gold Accent 4, Lighter 60%

                Dim xlStyleNegative As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleNegative.Font.FontColor = Color.Red

                Dim xlStyleDetail As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDetail.Font.FontColor = Color.Gray
                xlStyleDetail.Font.Italic = True
                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlPaxStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlPaxStyle.FormatCode = "#,##0;-#,##0;"

                If mReport.BooleanOption1 = 0 Then
                    .SetCellValue(1, 2, "INCLUDING OMIT")
                    .SetCellStyle(1, 2, xlStyleOMIT)
                End If
                .SetCellValue(2, 2, Format(mReport.Date1From, "dd/MM/yyyy") & " - " & Format(mReport.Date1To, "dd/MM/yyyy"))
                .SetCellValue(1, 5, "Total")
                .MergeWorksheetCells(1, 5, 1, 8)
                .SetCellValue(1, 9, "Total Air")
                .MergeWorksheetCells(1, 9, 1, 12)
                .SetCellValue(1, 13, "Total Services")
                .MergeWorksheetCells(1, 13, 1, 16)
                .SetCellValue(1, 17, "AIR")
                .MergeWorksheetCells(1, 17, 1, 36)
                .SetCellValue(1, 37, "SERVICES")
                .MergeWorksheetCells(1, 37, 1, 59)
                For ii = 0 To 1
                    .SetCellValue(2, ii * 20 + 17, "01 Marine")
                    .MergeWorksheetCells(2, ii * 20 + 17, 2, ii * 20 + 20)
                    .SetCellValue(2, ii * 20 + 21, "02 Interoffice")
                    .MergeWorksheetCells(2, ii * 20 + 21, 2, ii * 20 + 24)
                    .SetCellValue(2, ii * 20 + 25, "03 Corporate")
                    .MergeWorksheetCells(2, ii * 20 + 25, 2, ii * 20 + 28)
                    .SetCellValue(2, ii * 20 + 29, "04 Non Marine")
                    .MergeWorksheetCells(2, ii * 20 + 29, 2, ii * 20 + 32)
                    .SetCellValue(2, ii * 20 + 33, "05 Care of")
                    .MergeWorksheetCells(2, ii * 20 + 33, 2, ii * 20 + 36)
                Next
                .SetCellValue(2, 57, "RINVA")
                .MergeWorksheetCells(2, 57, 2, 59)
                .SetCellValue(5, 1, "Tots")
                .SetCellValue(5, 2, "Client Group")
                .SetCellValue(5, 3, "Client Code")
                .SetCellValue(5, 4, "Client Name")
                For ii = 0 To 12
                    .SetCellValue(5, ii * 4 + 5, "Net Payable")
                    .SetCellValue(5, ii * 4 + 6, "Net Buy")
                    .SetCellValue(5, ii * 4 + 7, "Profit")
                    .SetCellValue(5, ii * 4 + 8, "Pax")
                    .SetColumnStyle(ii * 4 + 5, ii * 4 + 7, xlDecStyle)
                    .SetColumnStyle(ii * 4 + 8, xlPaxStyle)
                Next
                .SetCellValue(5, 57, "Net Payable")
                .SetCellValue(5, 58, "Net Buy")
                .SetCellValue(5, 59, "Profit")
                .SetColumnStyle(57, 59, xlDecStyle)

                .SetCellStyle(1, 1, 5, 59, xlStyleHeader)
                xlINVCount = 5
                Dim Totals(59) As Decimal
                Dim TotClientGroup(59) As Decimal
                Dim TotsLevel As Integer
                Dim PrevClientGroup As String = ""
                Dim FirstGroupRow As Integer = 0
                Dim LastGroupRow As Integer = 0
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Dim Category As Integer = 0
                    Select Case mdsDataSet.Tables(0).Rows(i).Item(2)
                        Case "01"
                            Category = 1
                        Case "02"
                            Category = 2
                        Case "03"
                            Category = 3
                        Case "04"
                            Category = 4
                        Case Else
                            Category = 5
                    End Select
                    TotsLevel = mdsDataSet.Tables(0).Rows(i).Item(0)

                    If TotsLevel = 2 Or (TotsLevel = 1 And mdsDataSet.Tables(0).Rows(i).Item(3) <> "") Then
                        If TotsLevel = 2 And PrevClientGroup <> mdsDataSet.Tables(0).Rows(i).Item(1) Then
                            If PrevClientGroup <> "" Then
                                xlINVCount += 1
                                .SetCellValue(xlINVCount, 2, PrevClientGroup)
                                For j = 5 To 59
                                    .SetCellValue(xlINVCount, j, TotClientGroup(j))
                                    TotClientGroup(j) = 0
                                Next
                            End If
                            .GroupRows(FirstGroupRow, LastGroupRow)
                            .CollapseRows(LastGroupRow + 1)
                            FirstGroupRow = xlINVCount + 1
                        End If
                        xlINVCount += 1
                        LastGroupRow = xlINVCount
                        If TotsLevel = 2 Then
                            .SetCellStyle(xlINVCount, 1, xlINVCount, 59, xlStyleDetail)
                        End If
                        .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                        .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item(1))
                        .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item(3))
                        .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item(4))
                        ' Total
                        .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item(5) + mdsDataSet.Tables(0).Rows(i).Item(18) + mdsDataSet.Tables(0).Rows(i).Item(86))
                        .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item(6) + mdsDataSet.Tables(0).Rows(i).Item(19) + mdsDataSet.Tables(0).Rows(i).Item(87))
                        .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item(15) + mdsDataSet.Tables(0).Rows(i).Item(28) + mdsDataSet.Tables(0).Rows(i).Item(88))
                        .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item(16) + mdsDataSet.Tables(0).Rows(i).Item(29))

                        'AIR
                        .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item(5))
                        .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item(6))
                        .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item(15))
                        .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item(16))

                        .SetCellValue(xlINVCount, Category * 4 + 13, mdsDataSet.Tables(0).Rows(i).Item(5))
                        .SetCellValue(xlINVCount, Category * 4 + 14, mdsDataSet.Tables(0).Rows(i).Item(6))
                        .SetCellValue(xlINVCount, Category * 4 + 15, mdsDataSet.Tables(0).Rows(i).Item(15))
                        .SetCellValue(xlINVCount, Category * 4 + 16, mdsDataSet.Tables(0).Rows(i).Item(16))
                        'Services
                        .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item(18))
                        .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item(19))
                        .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item(28))
                        .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item(29))

                        .SetCellValue(xlINVCount, Category * 4 + 33, mdsDataSet.Tables(0).Rows(i).Item(18))
                        .SetCellValue(xlINVCount, Category * 4 + 34, mdsDataSet.Tables(0).Rows(i).Item(19))
                        .SetCellValue(xlINVCount, Category * 4 + 35, mdsDataSet.Tables(0).Rows(i).Item(28))
                        .SetCellValue(xlINVCount, Category * 4 + 36, mdsDataSet.Tables(0).Rows(i).Item(29))

                        ' RINVA
                        .SetCellValue(xlINVCount, 57, mdsDataSet.Tables(0).Rows(i).Item(86))
                        .SetCellValue(xlINVCount, 58, mdsDataSet.Tables(0).Rows(i).Item(87))
                        .SetCellValue(xlINVCount, 59, mdsDataSet.Tables(0).Rows(i).Item(88))
                        ' TOTALS
                        Totals(5) += mdsDataSet.Tables(0).Rows(i).Item(5) + mdsDataSet.Tables(0).Rows(i).Item(18) + mdsDataSet.Tables(0).Rows(i).Item(86)
                        Totals(6) += mdsDataSet.Tables(0).Rows(i).Item(6) + mdsDataSet.Tables(0).Rows(i).Item(19) + mdsDataSet.Tables(0).Rows(i).Item(87)
                        Totals(7) += mdsDataSet.Tables(0).Rows(i).Item(15) + mdsDataSet.Tables(0).Rows(i).Item(28) + mdsDataSet.Tables(0).Rows(i).Item(88)
                        Totals(8) += mdsDataSet.Tables(0).Rows(i).Item(16) + mdsDataSet.Tables(0).Rows(i).Item(29)
                        ' AIR
                        Totals(9) += mdsDataSet.Tables(0).Rows(i).Item(5)
                        Totals(10) += mdsDataSet.Tables(0).Rows(i).Item(6)
                        Totals(11) += mdsDataSet.Tables(0).Rows(i).Item(15)
                        Totals(12) += mdsDataSet.Tables(0).Rows(i).Item(16)
                        Totals(Category * 4 + 13) += mdsDataSet.Tables(0).Rows(i).Item(5)
                        Totals(Category * 4 + 14) += mdsDataSet.Tables(0).Rows(i).Item(6)
                        Totals(Category * 4 + 15) += mdsDataSet.Tables(0).Rows(i).Item(15)
                        Totals(Category * 4 + 16) += mdsDataSet.Tables(0).Rows(i).Item(16)
                        ' SERVICES
                        Totals(13) += mdsDataSet.Tables(0).Rows(i).Item(18)
                        Totals(14) += mdsDataSet.Tables(0).Rows(i).Item(19)
                        Totals(15) += mdsDataSet.Tables(0).Rows(i).Item(28)
                        Totals(16) += mdsDataSet.Tables(0).Rows(i).Item(29)

                        Totals(Category * 4 + 33) += mdsDataSet.Tables(0).Rows(i).Item(18)
                        Totals(Category * 4 + 34) += mdsDataSet.Tables(0).Rows(i).Item(19)
                        Totals(Category * 4 + 35) += mdsDataSet.Tables(0).Rows(i).Item(28)
                        Totals(Category * 4 + 36) += mdsDataSet.Tables(0).Rows(i).Item(29)
                        ' RINVA
                        Totals(57) += mdsDataSet.Tables(0).Rows(i).Item(86)
                        Totals(58) += mdsDataSet.Tables(0).Rows(i).Item(87)
                        Totals(59) += mdsDataSet.Tables(0).Rows(i).Item(88)

                        If TotsLevel = 2 Then

                            ' TOTALS
                            TotClientGroup(5) += mdsDataSet.Tables(0).Rows(i).Item(5) + mdsDataSet.Tables(0).Rows(i).Item(18) + mdsDataSet.Tables(0).Rows(i).Item(86)
                            TotClientGroup(6) += mdsDataSet.Tables(0).Rows(i).Item(6) + mdsDataSet.Tables(0).Rows(i).Item(19) + mdsDataSet.Tables(0).Rows(i).Item(87)
                            TotClientGroup(7) += mdsDataSet.Tables(0).Rows(i).Item(15) + mdsDataSet.Tables(0).Rows(i).Item(28) + mdsDataSet.Tables(0).Rows(i).Item(88)
                            TotClientGroup(8) += mdsDataSet.Tables(0).Rows(i).Item(16) + mdsDataSet.Tables(0).Rows(i).Item(29)
                            ' AIR
                            TotClientGroup(9) += mdsDataSet.Tables(0).Rows(i).Item(5)
                            TotClientGroup(10) += mdsDataSet.Tables(0).Rows(i).Item(6)
                            TotClientGroup(11) += mdsDataSet.Tables(0).Rows(i).Item(15)
                            TotClientGroup(12) += mdsDataSet.Tables(0).Rows(i).Item(16)
                            TotClientGroup(Category * 4 + 13) += mdsDataSet.Tables(0).Rows(i).Item(5)
                            TotClientGroup(Category * 4 + 14) += mdsDataSet.Tables(0).Rows(i).Item(6)
                            TotClientGroup(Category * 4 + 15) += mdsDataSet.Tables(0).Rows(i).Item(15)
                            TotClientGroup(Category * 4 + 16) += mdsDataSet.Tables(0).Rows(i).Item(16)
                            ' SERVICES
                            TotClientGroup(13) += mdsDataSet.Tables(0).Rows(i).Item(18)
                            TotClientGroup(14) += mdsDataSet.Tables(0).Rows(i).Item(19)
                            TotClientGroup(15) += mdsDataSet.Tables(0).Rows(i).Item(28)
                            TotClientGroup(16) += mdsDataSet.Tables(0).Rows(i).Item(29)

                            TotClientGroup(Category * 4 + 33) += mdsDataSet.Tables(0).Rows(i).Item(18)
                            TotClientGroup(Category * 4 + 34) += mdsDataSet.Tables(0).Rows(i).Item(19)
                            TotClientGroup(Category * 4 + 35) += mdsDataSet.Tables(0).Rows(i).Item(28)
                            TotClientGroup(Category * 4 + 36) += mdsDataSet.Tables(0).Rows(i).Item(29)
                            ' RINVA
                            TotClientGroup(57) += mdsDataSet.Tables(0).Rows(i).Item(86)
                            TotClientGroup(58) += mdsDataSet.Tables(0).Rows(i).Item(87)
                            TotClientGroup(59) += mdsDataSet.Tables(0).Rows(i).Item(88)

                            PrevClientGroup = mdsDataSet.Tables(0).Rows(i).Item(1)

                        End If
                    End If

                Next
                xlINVCount += 1
                .SetCellValue(xlINVCount, 2, PrevClientGroup)
                For j = 5 To 59
                    .SetCellValue(xlINVCount, j, TotClientGroup(j))
                    TotClientGroup(j) = 0
                Next
                .GroupRows(FirstGroupRow, LastGroupRow)
                .CollapseRows(LastGroupRow + 1)
                For i = 5 To 59
                    .SetCellValue(3, i, Totals(i))
                Next
                .SetCellStyle(1, 5, xlINVCount, 8, xlStyleTotalCol)
                .SetCellStyle(1, 9, xlINVCount, 12, xlStyleAir1)
                .SetCellStyle(1, 13, xlINVCount, 16, xlStyleServices1)
                .SetCellStyle(1, 17, xlINVCount, 20, xlStyleAir1)
                .SetCellStyle(1, 21, xlINVCount, 24, xlStyleAir2)
                .SetCellStyle(1, 25, xlINVCount, 28, xlStyleAir1)
                .SetCellStyle(1, 29, xlINVCount, 32, xlStyleAir2)
                .SetCellStyle(1, 33, xlINVCount, 36, xlStyleAir1)
                .SetCellStyle(1, 37, xlINVCount, 40, xlStyleServices1)
                .SetCellStyle(1, 41, xlINVCount, 44, xlStyleServices2)
                .SetCellStyle(1, 45, xlINVCount, 48, xlStyleServices1)
                .SetCellStyle(1, 49, xlINVCount, 52, xlStyleServices2)
                .SetCellStyle(1, 53, xlINVCount, 56, xlStyleServices1)
                .SetCellStyle(1, 57, xlINVCount, 59, xlStyleServices2)

                .DrawBorderGrid(1, 5, xlINVCount, 59, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .DrawBorderGrid(5, 1, xlINVCount, 4, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .AutoFitColumn(1, 59)
                .AddWorksheet("Summary")
                .SetCellValue(1, 3, "Net Payable")
                .SetCellValue(1, 4, "Net Buy")
                .SetCellValue(1, 5, "Profit")
                .SetCellValue(1, 6, "Pax")

                .SetCellValue(2, 1, "Total")

                .SetCellValue(4, 1, "AIR")
                .MergeWorksheetCells(4, 1, 9, 1)
                .SetCellValue(11, 1, "SERVICES")
                .MergeWorksheetCells(11, 1, 16, 1)
                For i = 4 To 11 Step 7
                    .SetCellValue(i, 2, "Total")
                    .SetCellValue(i + 1, 2, "01 Marine")
                    .SetCellValue(i + 2, 2, "02 Interoffice")
                    .SetCellValue(i + 3, 2, "03 Corporate")
                    .SetCellValue(i + 4, 2, "04 Non marine")
                    .SetCellValue(i + 5, 2, "05 Care of")
                Next
                .SetCellValue(18, 2, "RINVA")

                .SetCellValue(2, 3, Totals(5))
                .SetCellValue(2, 4, Totals(6))
                .SetCellValue(2, 5, Totals(7))
                .SetCellValue(2, 6, Totals(8))

                .SetCellValue(4, 3, Totals(9))
                .SetCellValue(4, 4, Totals(10))
                .SetCellValue(4, 5, Totals(11))
                .SetCellValue(4, 6, Totals(12))

                .SetCellValue(11, 3, Totals(13))
                .SetCellValue(11, 4, Totals(14))
                .SetCellValue(11, 5, Totals(15))
                .SetCellValue(11, 6, Totals(16))

                For j = 0 To 1
                    For i = 0 To 4
                        .SetCellValue(j * 7 + i + 5, 3, Totals(j * 20 + i * 4 + 17))
                        .SetCellValue(j * 7 + i + 5, 4, Totals(j * 20 + i * 4 + 18))
                        .SetCellValue(j * 7 + i + 5, 5, Totals(j * 20 + i * 4 + 19))
                        .SetCellValue(j * 7 + i + 5, 6, Totals(j * 20 + i * 4 + 20))
                    Next
                Next

                .SetCellValue(18, 3, Totals(57))
                .SetCellValue(18, 4, Totals(58))
                .SetCellValue(18, 5, Totals(59))

                .SetColumnStyle(3, 5, xlDecStyle)
                .SetColumnStyle(6, xlPaxStyle)
                'TOTALS
                xlStyleTotalCol.Font.Bold = True
                .SetCellStyle(1, 1, 2, 6, xlStyleTotalCol)
                .SetCellStyle(4, 2, 4, 6, xlStyleTotalCol)
                .SetCellStyle(11, 2, 11, 6, xlStyleTotalCol)
                'AIR
                .SetCellStyle(5, 2, 5, 6, xlStyleAir1)
                .SetCellStyle(6, 2, 6, 6, xlStyleAir2)
                .SetCellStyle(7, 2, 7, 6, xlStyleAir1)
                .SetCellStyle(8, 2, 8, 6, xlStyleAir2)
                .SetCellStyle(9, 2, 9, 6, xlStyleAir1)
                xlStyleAir1.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center)
                .SetCellStyle(4, 1, 4, 1, xlStyleAir1)
                'SERVICES
                .SetCellStyle(12, 2, 12, 6, xlStyleServices1)
                .SetCellStyle(13, 2, 13, 6, xlStyleServices2)
                .SetCellStyle(14, 2, 14, 6, xlStyleServices1)
                .SetCellStyle(15, 2, 15, 6, xlStyleServices2)
                .SetCellStyle(16, 2, 16, 6, xlStyleServices1)
                xlStyleServices1.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center)
                .SetCellStyle(11, 1, 11, 1, xlStyleServices1)
                'RINVA
                .SetCellStyle(18, 1, 18, 6, xlStyleServices2)

                .DrawBorderGrid(1, 1, 2, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .DrawBorderGrid(4, 1, 9, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .DrawBorderGrid(11, 1, 16, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .DrawBorderGrid(18, 1, 18, 6, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin)
                .AutoFitColumn(1, 6)

            End With


            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub
    Public Sub E52_Gaslog_Monthly_Statement()

    End Sub
    Public Sub E53_SeaChefs_InvoicesByDepartureDate(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim RowCounter As Integer = 0

        Dim Client As String = ""
        Dim ClientName As String = ""
        Dim Vessel As String = ""
        Dim CC As String = ""

        Dim PrevClient As String = ""
        Dim PrevVessel As String = ""
        Dim PrevCC As String = ""
        Dim Totals(4, 4) As Decimal

        Try
            With xlWorkSheet

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                xlStyleHeader.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlStyleHeader.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleDates As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDates.Font.Bold = True
                xlStyleDates.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleDates.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleDates.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left)

                Dim xlBold As SpreadsheetLight.SLStyle = .CreateStyle
                xlBold.Font.Bold = True

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 18, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(9, 13, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(4, xlNumStyle)

                RowCounter = 1


                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Client = mdsDataSet.Tables(0).Rows(i).Item(0)
                    ClientName = mdsDataSet.Tables(0).Rows(i).Item(1)
                    Vessel = mdsDataSet.Tables(0).Rows(i).Item(2)
                    CC = mdsDataSet.Tables(0).Rows(i).Item(3)
                    If RowCounter = 1 Or Client <> PrevClient Or Vessel <> PrevVessel Or CC <> PrevCC Then
                        If RowCounter > 1 Then
                            ' CC Total
                            RowCounter += 1
                            .SetCellValue(RowCounter, 7, PrevCC)
                            .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                            .SetCellValue(RowCounter, 9, Totals(0, 0))
                            .SetCellValue(RowCounter, 10, Totals(0, 1))
                            .SetCellValue(RowCounter, 11, Totals(0, 2))
                            .SetCellValue(RowCounter, 12, Totals(0, 3))
                            .SetCellValue(RowCounter, 13, Totals(0, 4))
                            .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)
                            Totals(0, 0) = 0
                            Totals(0, 1) = 0
                            Totals(0, 2) = 0
                            Totals(0, 3) = 0
                            Totals(0, 4) = 0

                            If Client <> PrevClient Or Vessel <> PrevVessel Then
                                ' Vessel Total
                                RowCounter += 1
                                .SetCellValue(RowCounter, 7, PrevVessel)
                                .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                                .SetCellValue(RowCounter, 9, Totals(1, 0))
                                .SetCellValue(RowCounter, 10, Totals(1, 1))
                                .SetCellValue(RowCounter, 11, Totals(1, 2))
                                .SetCellValue(RowCounter, 12, Totals(1, 3))
                                .SetCellValue(RowCounter, 13, Totals(1, 4))
                                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)
                                Totals(1, 0) = 0
                                Totals(1, 1) = 0
                                Totals(1, 2) = 0
                                Totals(1, 3) = 0
                                Totals(1, 4) = 0
                            End If
                            If Client <> PrevClient Then
                                ' Client Total
                                RowCounter += 1
                                .SetCellValue(RowCounter, 7, PrevClient)
                                .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                                .SetCellValue(RowCounter, 9, Totals(2, 0))
                                .SetCellValue(RowCounter, 10, Totals(2, 1))
                                .SetCellValue(RowCounter, 11, Totals(2, 2))
                                .SetCellValue(RowCounter, 12, Totals(2, 3))
                                .SetCellValue(RowCounter, 13, Totals(2, 4))
                                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)
                                Totals(2, 0) = 0
                                Totals(2, 1) = 0
                                Totals(2, 2) = 0
                                Totals(2, 3) = 0
                                Totals(2, 4) = 0
                            End If
                        End If
                        If RowCounter = 1 Or Client <> PrevClient Then
                            'If RowCounter = 1 Then RowCounter = 10

                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Period:({mReport.Date1From:dd/MM/yyyy} - {mReport.Date1To:dd/MM/yyyy})")
                            .MergeWorksheetCells(RowCounter, 1, RowCounter, 18)
                            .SetCellStyle(RowCounter, 1, xlStyleDates)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, "Currency")
                            .SetCellValue(RowCounter, 2, mdsDataSet.Tables(0).Rows(i).Item(7))
                            .SetCellValue(RowCounter, 3, "Account")
                            .SetCellValue(RowCounter, 4, Client)
                            .SetCellValue(RowCounter, 5, ClientName)
                            .SetCellStyle(RowCounter, 1, RowCounter, 5, xlBold)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, "Project")
                            .SetCellValue(RowCounter, 2, "Cost Center")
                            .SetCellValue(RowCounter, 3, "Invoice")
                            .SetCellValue(RowCounter, 4, "Dep Date")
                            .SetCellValue(RowCounter, 5, "AL")
                            .SetCellValue(RowCounter, 6, "Traveller")
                            .SetCellValue(RowCounter, 7, "TKT No")
                            .SetCellValue(RowCounter, 8, "Routing")
                            .SetCellValue(RowCounter, 9, "Fare")
                            .SetCellValue(RowCounter, 10, "Taxes")
                            .SetCellValue(RowCounter, 11, "Canc Fee")
                            .SetCellValue(RowCounter, 12, "Discount")
                            .SetCellValue(RowCounter, 13, "Payable")
                            .SetCellValue(RowCounter, 14, "Trip ID")
                            .SetCellValue(RowCounter, 15, "Order by")
                            .SetCellValue(RowCounter, 16, "Invoice Ref.")
                            .SetCellValue(RowCounter, 17, "Reason For travel")
                            .SetCellValue(RowCounter, 18, "Passenger ID")

                            .SetCellStyle(RowCounter, 1, RowCounter, 18, xlStyleHeader)
                        End If
                        If PrevClient <> Client Or PrevVessel <> Vessel Then
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, Vessel)
                            .SetCellStyle(RowCounter, 1, RowCounter, 1, xlBold)
                        End If
                        If PrevClient <> Client Or PrevVessel <> Vessel Or PrevCC <> CC Then
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, CC)
                            .SetCellStyle(RowCounter, 1, RowCounter, 1, xlBold)
                        End If

                        PrevClient = Client
                        PrevVessel = Vessel
                        PrevCC = CC
                    End If
                    RowCounter += 1
                    .SetCellValue(RowCounter, 1, mdsDataSet.Tables(0).Rows(i).Item(2))
                    .SetCellValue(RowCounter, 2, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(RowCounter, 3, mdsDataSet.Tables(0).Rows(i).Item(6))
                    .SetCellValue(RowCounter, 4, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(RowCounter, 5, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(RowCounter, 6, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(RowCounter, 7, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(RowCounter, 8, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(RowCounter, 9, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(RowCounter, 10, mdsDataSet.Tables(0).Rows(i).Item(14))
                    .SetCellValue(RowCounter, 11, mdsDataSet.Tables(0).Rows(i).Item(15))
                    .SetCellValue(RowCounter, 12, mdsDataSet.Tables(0).Rows(i).Item(16))
                    .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(RowCounter, 14, mdsDataSet.Tables(0).Rows(i).Item(18))
                    .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(19))
                    .SetCellValue(RowCounter, 16, "")
                    .SetCellValue(RowCounter, 17, mdsDataSet.Tables(0).Rows(i).Item(20))
                    .SetCellValue(RowCounter, 18, mdsDataSet.Tables(0).Rows(i).Item(21))
                    For iTot = 0 To 3
                        Totals(iTot, 0) += mdsDataSet.Tables(0).Rows(i).Item(13)
                        Totals(iTot, 1) += mdsDataSet.Tables(0).Rows(i).Item(14)
                        Totals(iTot, 2) += mdsDataSet.Tables(0).Rows(i).Item(15)
                        Totals(iTot, 3) += mdsDataSet.Tables(0).Rows(i).Item(16)
                        Totals(iTot, 4) += mdsDataSet.Tables(0).Rows(i).Item(17)
                    Next
                Next
                ' CC Total
                RowCounter += 1
                .SetCellValue(RowCounter, 7, PrevCC)
                .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                .SetCellValue(RowCounter, 9, Totals(0, 0))
                .SetCellValue(RowCounter, 10, Totals(0, 1))
                .SetCellValue(RowCounter, 11, Totals(0, 2))
                .SetCellValue(RowCounter, 12, Totals(0, 3))
                .SetCellValue(RowCounter, 13, Totals(0, 4))
                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)

                ' Vessel Total
                RowCounter += 1
                .SetCellValue(RowCounter, 7, PrevVessel)
                .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                .SetCellValue(RowCounter, 9, Totals(1, 0))
                .SetCellValue(RowCounter, 10, Totals(1, 1))
                .SetCellValue(RowCounter, 11, Totals(1, 2))
                .SetCellValue(RowCounter, 12, Totals(1, 3))
                .SetCellValue(RowCounter, 13, Totals(1, 4))
                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)

                ' Client Total
                RowCounter += 1
                .SetCellValue(RowCounter, 7, PrevClient)
                .SetCellValue(RowCounter, 8, "SUBTOTALS:")
                .SetCellValue(RowCounter, 9, Totals(2, 0))
                .SetCellValue(RowCounter, 10, Totals(2, 1))
                .SetCellValue(RowCounter, 11, Totals(2, 2))
                .SetCellValue(RowCounter, 12, Totals(2, 3))
                .SetCellValue(RowCounter, 13, Totals(2, 4))
                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)

                RowCounter += 1
                .SetCellValue(RowCounter, 8, "TOTALS:")
                .SetCellValue(RowCounter, 9, Totals(3, 0))
                .SetCellValue(RowCounter, 10, Totals(3, 1))
                .SetCellValue(RowCounter, 11, Totals(3, 2))
                .SetCellValue(RowCounter, 12, Totals(3, 3))
                .SetCellValue(RowCounter, 13, Totals(3, 4))
                .SetCellStyle(RowCounter, 7, RowCounter, 13, xlBold)

                .AutoFitColumn(1, 18)


            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E54_Client_Statement(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim RowCounter As Integer = 0

        Dim Client As String = ""
        Dim ClientName As String = ""
        Dim Vessel As String = ""

        Dim PrevClient As String = ""
        Dim PrevClientName As String = ""
        Dim PrevVessel As String = ""
        Dim Totals(3, 3) As Decimal

        Dim NumCols As Integer = 13
        Try
            With xlWorkSheet

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                xlStyleHeader.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlStyleHeader.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleDates As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDates.Font.Bold = True
                xlStyleDates.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleDates.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleDates.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left)

                Dim xlBold As SpreadsheetLight.SLStyle = .CreateStyle
                xlBold.Font.Bold = True
                xlBold.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlBold.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOmitted As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmitted.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmitted.Fill.SetPatternForegroundColor(Color.SandyBrown)

                If mReport.BooleanOption1 Then
                    .SetCellValue(1, 2, "INCLUDING OMIT")
                    .SetCellStyle(1, 2, xlStyleOMIT)
                    NumCols = 16
                End If
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, NumCols, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(10, 13, xlNumStyle)
                .SetColumnStyle(10, 16, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(2, xlNumStyle)
                .SetColumnStyle(8, xlNumStyle)

                RowCounter = 2
                .SetCellValue(RowCounter, 1, "ATPI Greece Travel Marine S.A. ,31-33 Athinon Avenue, 104 47 Athens, Greece")
                RowCounter += 1
                .SetCellValue(RowCounter, 1, "ATPI Ελλάς - Ταξειδιωτική Ναυτιλιακή Α.Ε., Λ.Αθηνών 31-33, 104 47, Αθήνα-ΑΦΜ 094333389 ΦΑΕ Αθηνών")
                RowCounter += 1

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Client = mdsDataSet.Tables(0).Rows(i).Item(0)
                    ClientName = mdsDataSet.Tables(0).Rows(i).Item(1)
                    Vessel = mdsDataSet.Tables(0).Rows(i).Item(5)
                    If RowCounter = 4 Or Client <> PrevClient Or Vessel <> PrevVessel Then
                        If RowCounter > 4 Then
                            ' Vessel Total
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                            .SetCellValue(RowCounter, 10, Totals(0, 0))
                            .SetCellValue(RowCounter, 11, Totals(0, 1))
                            .SetCellValue(RowCounter, 12, Totals(0, 2))
                            .SetCellValue(RowCounter, 13, Totals(0, 3))
                            .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlBold)
                            Totals(0, 0) = 0
                            Totals(0, 1) = 0
                            Totals(0, 2) = 0
                            Totals(0, 3) = 0

                            If Client <> PrevClient Then
                                ' Client Total
                                RowCounter += 1
                                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                                .SetCellValue(RowCounter, 10, Totals(1, 0))
                                .SetCellValue(RowCounter, 11, Totals(1, 1))
                                .SetCellValue(RowCounter, 12, Totals(1, 2))
                                .SetCellValue(RowCounter, 13, Totals(1, 3))
                                .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlBold)
                                Totals(1, 0) = 0
                                Totals(1, 1) = 0
                                Totals(1, 2) = 0
                                Totals(1, 3) = 0
                            End If
                        End If
                        If RowCounter = 1 Or Client <> PrevClient Then

                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Statement Period:({mReport.Date1From:dd/MM/yyyy} - {mReport.Date1To:dd/MM/yyyy})")
                            .MergeWorksheetCells(RowCounter, 1, RowCounter, NumCols)
                            .SetCellStyle(RowCounter, 1, xlStyleDates)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Client: {Client} {ClientName}")
                            .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(2))
                            .SetCellStyle(RowCounter, 1, RowCounter, 5, xlBold)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, "P-T Number")
                            .SetCellValue(RowCounter, 2, "Inv. Date")
                            .SetCellValue(RowCounter, 3, "Vessel")
                            .SetCellValue(RowCounter, 4, "Description")
                            .SetCellValue(RowCounter, 5, "Pax")
                            .SetCellValue(RowCounter, 6, "A/L")
                            .SetCellValue(RowCounter, 7, "Routing")
                            .SetCellValue(RowCounter, 8, "Flight Date")
                            .SetCellValue(RowCounter, 9, "ReasonForTravel")
                            .SetCellValue(RowCounter, 10, "Taxes")
                            .SetCellValue(RowCounter, 11, "Face Value")
                            .SetCellValue(RowCounter, 12, "Discount")
                            .SetCellValue(RowCounter, 13, "Net Payable")
                            If mReport.BooleanOption1 Then
                                .SetCellValue(RowCounter, 14, "OMIT")
                                .SetCellValue(RowCounter, 15, "ConnectedDocument")
                                .SetCellValue(RowCounter, 16, "ConnectedDoc.Amount")
                            End If
                            .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlStyleHeader)
                        End If
                        If PrevClient <> Client Or PrevVessel <> Vessel Then
                            RowCounter += 1
                            '.SetCellValue(RowCounter, 1, Vessel)
                            '.SetCellStyle(RowCounter, 1, RowCounter, 1, xlBold)
                        End If

                        PrevClient = Client
                        PrevClientName = ClientName
                        PrevVessel = Vessel
                    End If
                    RowCounter += 1
                    .SetCellValue(RowCounter, 1, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(RowCounter, 2, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(RowCounter, 3, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(RowCounter, 4, mdsDataSet.Tables(0).Rows(i).Item(6))
                    .SetCellValue(RowCounter, 5, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(RowCounter, 6, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(RowCounter, 7, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(RowCounter, 8, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(RowCounter, 9, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(RowCounter, 10, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(RowCounter, 11, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(RowCounter, 12, mdsDataSet.Tables(0).Rows(i).Item(14))
                    .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(15))
                    If mReport.BooleanOption1 Then
                        If mdsDataSet.Tables(0).Rows(i).Item(16) Then
                            .SetCellValue(RowCounter, 14, "OMIT")
                            .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlStyleOmitted)
                        End If
                        .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(17))
                        .SetCellValue(RowCounter, 16, mdsDataSet.Tables(0).Rows(i).Item(18))
                    End If
                    For iTot = 0 To 2
                        Totals(iTot, 0) += mdsDataSet.Tables(0).Rows(i).Item(12)
                        Totals(iTot, 1) += mdsDataSet.Tables(0).Rows(i).Item(13)
                        Totals(iTot, 2) += mdsDataSet.Tables(0).Rows(i).Item(14)
                        Totals(iTot, 3) += mdsDataSet.Tables(0).Rows(i).Item(15)
                    Next
                Next

                ' Vessel Total
                RowCounter += 1
                .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                .SetCellValue(RowCounter, 10, Totals(0, 0))
                .SetCellValue(RowCounter, 11, Totals(0, 1))
                .SetCellValue(RowCounter, 12, Totals(0, 2))
                .SetCellValue(RowCounter, 13, Totals(0, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlBold)

                ' Client Total
                RowCounter += 1
                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                .SetCellValue(RowCounter, 10, Totals(1, 0))
                .SetCellValue(RowCounter, 11, Totals(1, 1))
                .SetCellValue(RowCounter, 12, Totals(1, 2))
                .SetCellValue(RowCounter, 13, Totals(1, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlBold)

                RowCounter += 1
                .SetCellValue(RowCounter, 1, "TOTALS:")
                .SetCellValue(RowCounter, 10, Totals(2, 0))
                .SetCellValue(RowCounter, 11, Totals(2, 1))
                .SetCellValue(RowCounter, 12, Totals(2, 2))
                .SetCellValue(RowCounter, 13, Totals(2, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, NumCols, xlBold)

                .AutoFitColumn(1, NumCols)


            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E55_Safety_Statement(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim RowCounter As Integer = 0

        Dim Client As String = ""
        Dim ClientName As String = ""
        Dim Vessel As String = ""

        Dim PrevClient As String = ""
        Dim PrevClientName As String = ""
        Dim PrevVessel As String = ""
        Dim Totals(3, 3) As Decimal

        Try
            With xlWorkSheet

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                xlStyleHeader.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlStyleHeader.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleDates As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDates.Font.Bold = True
                xlStyleDates.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleDates.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleDates.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left)

                Dim xlBold As SpreadsheetLight.SLStyle = .CreateStyle
                xlBold.Font.Bold = True
                xlBold.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlBold.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlClient As SpreadsheetLight.SLStyle = .CreateStyle
                xlClient.Font.Bold = True
                xlClient.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlClient.Fill.SetPatternForegroundColor(Color.Cyan)
                xlClient.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlClient.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOmitted As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmitted.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmitted.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim ColumnCount As Integer = 15
                If mReport.BooleanOption1 Then ' With Fare Basis
                    ColumnCount = 18
                End If
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, ColumnCount, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(12, 15, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(2, xlNumStyle)
                .SetColumnStyle(11, xlNumStyle)

                RowCounter = 2
                .SetCellValue(RowCounter, 1, "ATPI Greece Travel Marine S.A. ,31-33 Athinon Avenue, 104 47 Athens, Greece")
                .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                RowCounter += 1
                .SetCellValue(RowCounter, 1, "ATPI Ελλάς - Ταξειδιωτική Ναυτιλιακή Α.Ε., Λ.Αθηνών 31-33, 104 47, Αθήνα-ΑΦΜ 094333389 ΦΑΕ Αθηνών")
                .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                RowCounter += 1

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Client = mdsDataSet.Tables(0).Rows(i).Item(0)
                    ClientName = mdsDataSet.Tables(0).Rows(i).Item(1)
                    Vessel = mdsDataSet.Tables(0).Rows(i).Item(5)
                    If RowCounter = 4 Or Client <> PrevClient Or Vessel <> PrevVessel Then
                        If RowCounter > 4 Then
                            ' Vessel Total
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                            .SetCellValue(RowCounter, 12, Totals(0, 0))
                            .SetCellValue(RowCounter, 13, Totals(0, 1))
                            .SetCellValue(RowCounter, 14, Totals(0, 2))
                            .SetCellValue(RowCounter, 15, Totals(0, 3))
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)
                            Totals(0, 0) = 0
                            Totals(0, 1) = 0
                            Totals(0, 2) = 0
                            Totals(0, 3) = 0

                            If Client <> PrevClient Then
                                ' Client Total
                                RowCounter += 1
                                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                                .SetCellValue(RowCounter, 12, Totals(1, 0))
                                .SetCellValue(RowCounter, 13, Totals(1, 1))
                                .SetCellValue(RowCounter, 14, Totals(1, 2))
                                .SetCellValue(RowCounter, 15, Totals(1, 3))
                                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)
                                Totals(1, 0) = 0
                                Totals(1, 1) = 0
                                Totals(1, 2) = 0
                                Totals(1, 3) = 0
                            End If
                        End If
                        If RowCounter = 1 Or Client <> PrevClient Then

                            RowCounter += 2
                            .SetCellValue(RowCounter, 1, $"Statement Period:({mReport.Date1From:dd/MM/yyyy} - {mReport.Date1To:dd/MM/yyyy})")
                            .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                            .SetCellStyle(RowCounter, 1, xlStyleDates)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Client: {Client} {ClientName}")
                            .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(2))
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlClient)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, "P-T Number")
                            .SetCellValue(RowCounter, 2, "Inv. Date")
                            .SetCellValue(RowCounter, 3, "Vessel")
                            .SetCellValue(RowCounter, 4, "Description")
                            .SetCellValue(RowCounter, 5, "Booking Class")
                            .SetCellValue(RowCounter, 6, "Class Of Service")
                            .SetCellValue(RowCounter, 7, "Tkt")

                            .SetCellValue(RowCounter, 8, "Pax")
                            .SetCellValue(RowCounter, 9, "A/L")
                            .SetCellValue(RowCounter, 10, "Routing")
                            .SetCellValue(RowCounter, 11, "Flight Date")
                            .SetCellValue(RowCounter, 12, "Fare")
                            .SetCellValue(RowCounter, 13, "Taxes")
                            .SetCellValue(RowCounter, 14, "Discount")
                            .SetCellValue(RowCounter, 15, "Total Payable")
                            If mReport.BooleanOption1 Then
                                .SetCellValue(RowCounter, 16, "Fare Basis")
                            End If
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlStyleHeader)
                        End If
                        If PrevClient <> Client Or PrevVessel <> Vessel Then
                            RowCounter += 1
                        End If

                        PrevClient = Client
                        PrevClientName = ClientName
                        PrevVessel = Vessel
                    End If
                    RowCounter += 1
                    .SetCellValue(RowCounter, 1, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(RowCounter, 2, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(RowCounter, 3, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(RowCounter, 4, mdsDataSet.Tables(0).Rows(i).Item(6))

                    .SetCellValue(RowCounter, 5, mdsDataSet.Tables(0).Rows(i).Item(16))
                    .SetCellValue(RowCounter, 6, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(RowCounter, 7, mdsDataSet.Tables(0).Rows(i).Item(18))

                    .SetCellValue(RowCounter, 8, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(RowCounter, 9, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(RowCounter, 10, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(RowCounter, 11, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(RowCounter, 12, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(RowCounter, 14, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(14))
                    If mReport.BooleanOption1 Then
                        .SetCellValue(RowCounter, 16, mdsDataSet.Tables(0).Rows(i).Item("FareBasis"))
                    End If
                    For iTot = 0 To 2
                        Totals(iTot, 0) += mdsDataSet.Tables(0).Rows(i).Item(12)
                        Totals(iTot, 1) += mdsDataSet.Tables(0).Rows(i).Item(11)
                        Totals(iTot, 2) += mdsDataSet.Tables(0).Rows(i).Item(13)
                        Totals(iTot, 3) += mdsDataSet.Tables(0).Rows(i).Item(14)
                    Next
                Next

                ' Vessel Total
                RowCounter += 1
                .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                .SetCellValue(RowCounter, 12, Totals(0, 0))
                .SetCellValue(RowCounter, 13, Totals(0, 1))
                .SetCellValue(RowCounter, 14, Totals(0, 2))
                .SetCellValue(RowCounter, 15, Totals(0, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                ' Client Total
                RowCounter += 1
                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                .SetCellValue(RowCounter, 12, Totals(1, 0))
                .SetCellValue(RowCounter, 13, Totals(1, 1))
                .SetCellValue(RowCounter, 14, Totals(1, 2))
                .SetCellValue(RowCounter, 15, Totals(1, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                RowCounter += 1
                .SetCellValue(RowCounter, 1, "TOTALS:")
                .SetCellValue(RowCounter, 12, Totals(2, 0))
                .SetCellValue(RowCounter, 13, Totals(2, 1))
                .SetCellValue(RowCounter, 14, Totals(2, 2))
                .SetCellValue(RowCounter, 15, Totals(2, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                .AutoFitColumn(1, ColumnCount)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E56_ClientsPerGroup(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim xlINVCount As Integer = 0

        Try
            With xlWorkSheet
                .FreezePanes(1, 0)


                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlTextStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlTextStyle.FormatCode = "@"
                .SetColumnStyle(1, 4, xlTextStyle)
                Dim xlDateStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlDateStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(5, 6, xlDateStyle)

                .SetCellValue(1, 1, "Client Code")
                .SetCellValue(1, 2, "Client Name")
                .SetCellValue(1, 3, "Client Group")
                .SetCellValue(1, 4, "Agent Group")
                .SetCellValue(1, 5, "Last Transaction")
                .SetCellValue(1, 6, "Date Created")

                .SetCellStyle(1, 1, 1, 6, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    For j = 0 To 5
                        If (Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(j))) Then
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))

                        End If
                    Next
                    If (IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(4)) OrElse DateDiff("m", mdsDataSet.Tables(0).Rows(i).Item(4), Today) > 24) Then
                        .SetCellStyle(xlINVCount, 5, xlStyleOmit)
                    End If
                Next

                .AutoFitColumn(1, 6)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E60_Report_For_Lowest_Class(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)

        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim RowCounter As Integer = 0

        Dim Client As String = ""
        Dim ClientName As String = ""
        Dim Vessel As String = ""

        Try
            With xlWorkSheet

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                xlStyleHeader.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlStyleHeader.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleDates As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDates.Font.Bold = True
                xlStyleDates.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleDates.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleDates.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left)

                Dim xlBold As SpreadsheetLight.SLStyle = .CreateStyle
                xlBold.Font.Bold = True
                xlBold.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlBold.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlClient As SpreadsheetLight.SLStyle = .CreateStyle
                xlClient.Font.Bold = True
                xlClient.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlClient.Fill.SetPatternForegroundColor(Color.Cyan)
                xlClient.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlClient.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOmitted As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmitted.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmitted.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim ColumnCount As Integer = 23

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, ColumnCount, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(15, 18, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(5, 6, xlNumStyle)
                .SetColumnStyle(14, xlNumStyle)
                .FreezePanes(1, 0)
                RowCounter = 1
                .SetCellValue(RowCounter, 1, "Client")
                .SetCellValue(RowCounter, 2, "Client Name")
                .SetCellValue(RowCounter, 3, "Vessel")
                .SetCellValue(RowCounter, 4, "P-T Number")
                .SetCellValue(RowCounter, 5, "Inv. Date")
                .SetCellValue(RowCounter, 6, "Booking Date")
                .SetCellValue(RowCounter, 7, "Description")
                .SetCellValue(RowCounter, 8, "Class Of Service")
                .SetCellValue(RowCounter, 9, "Cabin")
                .SetCellValue(RowCounter, 10, "Tkt")

                .SetCellValue(RowCounter, 11, "Pax")
                .SetCellValue(RowCounter, 12, "A/L")
                .SetCellValue(RowCounter, 13, "Routing")
                .SetCellValue(RowCounter, 14, "Flight Date")
                .SetCellValue(RowCounter, 15, "Fare")
                .SetCellValue(RowCounter, 16, "Taxes")
                .SetCellValue(RowCounter, 17, "Discount")
                .SetCellValue(RowCounter, 18, "Total Payable")
                .SetCellValue(RowCounter, 19, "Fare Basis")
                .SetCellValue(RowCounter, 20, "Transaction Type")
                .SetCellValue(RowCounter, 21, "VOID")
                .SetCellValue(RowCounter, 22, "Connected Document")
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlStyleHeader)

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Client = mdsDataSet.Tables(0).Rows(i).Item(0)
                    ClientName = mdsDataSet.Tables(0).Rows(i).Item(1)
                    Vessel = mdsDataSet.Tables(0).Rows(i).Item(5)

                    RowCounter += 1
                    .SetCellValue(RowCounter, 1, Client)
                    .SetCellValue(RowCounter, 2, ClientName)
                    .SetCellValue(RowCounter, 3, Vessel)
                    .SetCellValue(RowCounter, 4, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(RowCounter, 5, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(RowCounter, 6, mdsDataSet.Tables(0).Rows(i).Item("EntryDate"))
                    .SetCellValue(RowCounter, 7, mdsDataSet.Tables(0).Rows(i).Item(6))

                    .SetCellValue(RowCounter, 8, mdsDataSet.Tables(0).Rows(i).Item(16))
                    .SetCellValue(RowCounter, 9, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(RowCounter, 10, mdsDataSet.Tables(0).Rows(i).Item(18))

                    .SetCellValue(RowCounter, 11, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(RowCounter, 12, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(RowCounter, 14, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(RowCounter, 16, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(RowCounter, 17, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(RowCounter, 18, mdsDataSet.Tables(0).Rows(i).Item(14))
                    .SetCellValue(RowCounter, 19, mdsDataSet.Tables(0).Rows(i).Item("FareBasis"))
                    .SetCellValue(RowCounter, 20, mdsDataSet.Tables(0).Rows(i).Item("TransactionType"))
                    .SetCellValue(RowCounter, 21, mdsDataSet.Tables(0).Rows(i).Item("Void"))
                    .SetCellValue(RowCounter, 22, mdsDataSet.Tables(0).Rows(i).Item("ConnectedDocument"))
                    If (mdsDataSet.Tables(0).Rows(i).Item("Omit") = 1) Then
                        .SetCellValue(RowCounter, 23, "OMIT")
                        .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlStyleOMIT)
                    End If
                Next
                .AutoFitColumn(1, ColumnCount)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E63_AirTicketSalesTemenos(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleGreyOut As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGreyOut.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGreyOut.Fill.SetPatternForegroundColor(Color.LightGray)

                Dim xlStyleMandatory As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleMandatory.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleMandatory.Fill.SetPatternForegroundColor(Color.Yellow)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 20, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(6, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(4, xlNumStyle)
                .SetColumnStyle(10, 11, xlNumStyle)

                .SetCellValue(1, 1, "Passenger") ' 1
                .SetCellValue(1, 2, "Inv Code") '2
                .SetCellValue(1, 3, "Inv Number") '3
                .SetCellValue(1, 4, "Invoice Date") '4
                .SetCellValue(1, 5, "Vessel") '5
                .SetCellValue(1, 6, "Net Payable") '6
                .SetCellValue(1, 7, "Transaction Type") '7
                .SetCellValue(1, 8, "Airline") '8
                .SetCellValue(1, 9, "Routing") '9
                .SetCellValue(1, 10, "Departure Date") '10
                .SetCellValue(1, 11, "Arrival Date") '11
                .SetCellValue(1, 12, "PNR") ' 12
                .SetCellValue(1, 13, "BookedBy") '13
                .SetCellValue(1, 14, "ReasonForTravel") '14
                .SetCellValue(1, 15, "CostCentre") '15
                .SetCellValue(1, 16, "Employee ID") '16
                .SetCellValue(1, 17, "Project LIST A") '17
                .SetCellValue(1, 18, "Employee Type") '18
                .SetCellValue(1, 19, "Entity Code") '19
                .SetCellValue(1, 20, "Project LIST B") '20

                .SetCellStyle(1, 1, 1, 20, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1

                    xlINVCount += 1

                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item("Passenger"))
                    If mdsDataSet.Tables(0).Rows(i).Item("InvNumber") <> 0 Then
                        .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item("InvCode"))
                        .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item("InvNumber"))
                        .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item("InvoiceDate"))
                    End If
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item("Vessel"))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item("NetPayable"))
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item("TransactionType"))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item("TicketingAirline"))
                    .SetCellValue(xlINVCount, 9, mdsDataSet.Tables(0).Rows(i).Item("Routing"))
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item("DepartureDate"))
                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item("ArrivalDate"))
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item("PNR"))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item("01-BookedBy"))
                    .SetCellValue(xlINVCount, 14, mdsDataSet.Tables(0).Rows(i).Item("04-ReasonForTravel"))
                    .SetCellValue(xlINVCount, 15, mdsDataSet.Tables(0).Rows(i).Item("05-CostCentre"))
                    .SetCellValue(xlINVCount, 16, mdsDataSet.Tables(0).Rows(i).Item("12-Passenger ID"))
                    .SetCellValue(xlINVCount, 17, mdsDataSet.Tables(0).Rows(i).Item("18-REF1"))
                    .SetCellValue(xlINVCount, 18, mdsDataSet.Tables(0).Rows(i).Item("20-REF3"))
                    .SetCellValue(xlINVCount, 19, mdsDataSet.Tables(0).Rows(i).Item("23-REF6"))
                    .SetCellValue(xlINVCount, 20, mdsDataSet.Tables(0).Rows(i).Item("19-REF2"))

                    If mdsDataSet.Tables(0).Rows(i).Item("Omit") <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 20, xlStyleOmit)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item("Void") <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 20, xlStyleVoid)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item("ActionType") = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 20, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 20)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E64_LowestClasses(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument
        Dim RowCounter As Integer = 0

        Dim Client As String = ""
        Dim ClientName As String = ""
        Dim Vessel As String = ""
        Dim BookingClass As String = ""
        Dim ClassRank As String = ""
        Dim ClassLevel As Integer
        Dim Itin As String = ""
        Dim AirTicketTypeId As Long
        Dim AirTicketType As String = ""
        Dim IsExchanged As Boolean
        Dim IsExchangedText As String = ""
        Dim HasRelatedTicket As Boolean
        Dim PrevClient As String = ""
        Dim PrevClientName As String = ""
        Dim PrevVessel As String = ""
        Dim Totals(3, 3) As Decimal

        Try
            With xlWorkSheet

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                xlStyleHeader.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlStyleHeader.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleDates As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleDates.Font.Bold = True
                xlStyleDates.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleDates.Fill.SetPatternForegroundColor(Color.FromArgb(255, 255, 224, 192))
                xlStyleDates.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left)

                Dim xlBold As SpreadsheetLight.SLStyle = .CreateStyle
                xlBold.Font.Bold = True
                xlBold.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlBold.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlClient As SpreadsheetLight.SLStyle = .CreateStyle
                xlClient.Font.Bold = True
                xlClient.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlClient.Fill.SetPatternForegroundColor(Color.Cyan)
                xlClient.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)
                xlClient.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, Color.Black)

                Dim xlStyleOMIT As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOMIT.Font.Bold = True
                xlStyleOMIT.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOMIT.Fill.SetPatternForegroundColor(Color.OrangeRed)
                xlStyleOMIT.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleOmitted As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmitted.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmitted.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleGreen As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGreen.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGreen.Fill.SetPatternForegroundColor(Color.LightGreen)

                Dim xlCentred As SpreadsheetLight.SLStyle = .CreateStyle
                xlCentred.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center
                Dim xlRoundTripStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlRoundTripStyle.Font.Bold = True
                xlRoundTripStyle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlRoundTripStyle.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow)

                Dim ColumnCount As Integer = 18
                If mReport.BooleanOption1 Then ' With Fare Basis
                    ColumnCount = 19
                End If
                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, ColumnCount, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(14, 16, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;#,##0.00"
                .SetColumnStyle(17, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(2, xlNumStyle)
                .SetColumnStyle(13, xlNumStyle)
                .SetColumnStyle(5, 6, xlCentred)
                RowCounter = 2
                .SetCellValue(RowCounter, 1, "ATPI Greece Travel Marine S.A., 31-33 Athinon Avenue, 104 47 Athens, Greece")
                .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                RowCounter += 1
                .SetCellValue(RowCounter, 1, "ATPI Ελλάς - Ταξειδιωτική Ναυτιλιακή Α.Ε., Λ.Αθηνών 31-33, 104 47, Αθήνα-ΑΦΜ 094333389 ΦΑΕ Αθηνών")
                .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                RowCounter += 1

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    Client = mdsDataSet.Tables(0).Rows(i).Item(0)
                    ClientName = mdsDataSet.Tables(0).Rows(i).Item(1)
                    Vessel = mdsDataSet.Tables(0).Rows(i).Item(5)
                    BookingClass = mdsDataSet.Tables(0).Rows(i).Item(16)
                    ClassRank = mdsDataSet.Tables(0).Rows(i).Item(26)
                    ClassLevel = GetClassLevel(BookingClass, ClassRank)
                    Itin = mdsDataSet.Tables(0).Rows(i).Item(9)
                    AirTicketTypeId = mdsDataSet.Tables(0).Rows(i).Item(19)
                    AirTicketType = mdsDataSet.Tables(0).Rows(i).Item(29)
                    IsExchanged = mdsDataSet.Tables(0).Rows(i).Item(31)
                    If IsExchanged Then
                        IsExchangedText = $"Reissue {mdsDataSet.Tables(0).Rows(i).Item(32)}"
                    Else
                        IsExchangedText = ""
                    End If
                    HasRelatedTicket = mdsDataSet.Tables(0).Rows(i).Item(33)
                    If HasRelatedTicket Then
                        If IsExchangedText <> "" Then
                            IsExchangedText &= " - "
                        End If
                        IsExchangedText &= $"Related {mdsDataSet.Tables(0).Rows(i).Item(34)}"
                    End If
                    If RowCounter = 4 Or Client <> PrevClient Or Vessel <> PrevVessel Then
                        If RowCounter > 4 Then
                            ' Vessel Total
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                            .SetCellValue(RowCounter, 14, Totals(0, 0))
                            .SetCellValue(RowCounter, 15, Totals(0, 1))
                            .SetCellValue(RowCounter, 16, Totals(0, 2))
                            .SetCellValue(RowCounter, 17, Totals(0, 3))
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)
                            Totals(0, 0) = 0
                            Totals(0, 1) = 0
                            Totals(0, 2) = 0
                            Totals(0, 3) = 0

                            If Client <> PrevClient Then
                                ' Client Total
                                RowCounter += 1
                                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                                .SetCellValue(RowCounter, 14, Totals(1, 0))
                                .SetCellValue(RowCounter, 15, Totals(1, 1))
                                .SetCellValue(RowCounter, 16, Totals(1, 2))
                                .SetCellValue(RowCounter, 17, Totals(1, 3))
                                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)
                                Totals(1, 0) = 0
                                Totals(1, 1) = 0
                                Totals(1, 2) = 0
                                Totals(1, 3) = 0
                            End If
                        End If
                        If RowCounter = 1 Or Client <> PrevClient Then

                            RowCounter += 2
                            .SetCellValue(RowCounter, 1, $"Statement Period:({mReport.Date1From:dd/MM/yyyy} - {mReport.Date1To:dd/MM/yyyy})")
                            .MergeWorksheetCells(RowCounter, 1, RowCounter, ColumnCount)
                            .SetCellStyle(RowCounter, 1, xlStyleDates)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, $"Client: {Client} {ClientName}")
                            .SetCellValue(RowCounter, 17, mdsDataSet.Tables(0).Rows(i).Item(2))
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlClient)
                            RowCounter += 1
                            .SetCellValue(RowCounter, 1, "P-T Number")
                            .SetCellValue(RowCounter, 2, "Inv. Date")
                            .SetCellValue(RowCounter, 3, "Vessel")
                            .SetCellValue(RowCounter, 4, "Description")
                            .SetCellValue(RowCounter, 5, "Booking Class")
                            .SetCellValue(RowCounter, 6, "Class Rank")
                            .SetCellValue(RowCounter, 7, "Marine")
                            .SetCellValue(RowCounter, 8, "Class Of Service")
                            .SetCellValue(RowCounter, 9, "Tkt")

                            .SetCellValue(RowCounter, 10, "Pax")
                            .SetCellValue(RowCounter, 11, "A/L")
                            .SetCellValue(RowCounter, 12, "Routing")
                            .SetCellValue(RowCounter, 13, "Flight Date")
                            .SetCellValue(RowCounter, 14, "Fare")
                            .SetCellValue(RowCounter, 15, "Taxes")
                            .SetCellValue(RowCounter, 16, "Discount")
                            .SetCellValue(RowCounter, 17, "Total Payable")
                            If mReport.BooleanOption1 Then
                                .SetCellValue(RowCounter, 18, "Fare Basis")
                                .SetCellValue(RowCounter, 19, "Notes")
                            Else
                                .SetCellValue(RowCounter, 18, "Notes")
                            End If
                            .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlStyleHeader)
                        End If
                        If PrevClient <> Client Or PrevVessel <> Vessel Then
                            RowCounter += 1
                        End If

                        PrevClient = Client
                        PrevClientName = ClientName
                        PrevVessel = Vessel
                    End If
                    RowCounter += 1

                    .SetCellValue(RowCounter, 1, mdsDataSet.Tables(0).Rows(i).Item(3))
                    .SetCellValue(RowCounter, 2, mdsDataSet.Tables(0).Rows(i).Item(4))
                    .SetCellValue(RowCounter, 3, mdsDataSet.Tables(0).Rows(i).Item(5))
                    .SetCellValue(RowCounter, 4, mdsDataSet.Tables(0).Rows(i).Item(6))

                    .SetCellValue(RowCounter, 5, BookingClass)

                    If AirTicketTypeId <> 323 Then
                        .SetCellValue(RowCounter, 6, AirTicketType)
                        .SetCellStyle(RowCounter, 6, xlRoundTripStyle)
                    Else
                        If Itin.Length > 6 Then
                            If Itin.Substring(0, 3) = Itin.Substring(Itin.Length - 3, 3) Then
                                .SetCellValue(RowCounter, 6, "RT")
                                .SetCellStyle(RowCounter, 6, xlRoundTripStyle)
                            Else
                                If ClassLevel > 0 Then
                                    .SetCellValue(RowCounter, 6, ClassLevel)
                                    If ClassLevel <= 2 Then
                                        .SetCellStyle(RowCounter, 6, xlStyleGreen)
                                    End If
                                End If
                            End If
                        End If
                        If mdsDataSet.Tables(0).Rows(i).Item(27) = 1 Then
                            .SetCellValue(RowCounter, 7, "Marine")
                        End If
                    End If


                    .SetCellValue(RowCounter, 8, mdsDataSet.Tables(0).Rows(i).Item(17))
                    .SetCellValue(RowCounter, 9, mdsDataSet.Tables(0).Rows(i).Item(18))

                    .SetCellValue(RowCounter, 10, mdsDataSet.Tables(0).Rows(i).Item(7))
                    .SetCellValue(RowCounter, 11, mdsDataSet.Tables(0).Rows(i).Item(8))
                    .SetCellValue(RowCounter, 12, mdsDataSet.Tables(0).Rows(i).Item(9))
                    .SetCellValue(RowCounter, 13, mdsDataSet.Tables(0).Rows(i).Item(10))
                    .SetCellValue(RowCounter, 14, mdsDataSet.Tables(0).Rows(i).Item(12))
                    .SetCellValue(RowCounter, 15, mdsDataSet.Tables(0).Rows(i).Item(11))
                    .SetCellValue(RowCounter, 16, mdsDataSet.Tables(0).Rows(i).Item(13))
                    .SetCellValue(RowCounter, 17, mdsDataSet.Tables(0).Rows(i).Item(14))
                    If mReport.BooleanOption1 Then
                        .SetCellValue(RowCounter, 18, mdsDataSet.Tables(0).Rows(i).Item("FareBasis"))
                        .SetCellValue(RowCounter, 19, IsExchangedText)
                    Else
                        .SetCellValue(RowCounter, 18, IsExchangedText)
                    End If
                    For iTot = 0 To 2
                        Totals(iTot, 0) += mdsDataSet.Tables(0).Rows(i).Item(12)
                        Totals(iTot, 1) += mdsDataSet.Tables(0).Rows(i).Item(11)
                        Totals(iTot, 2) += mdsDataSet.Tables(0).Rows(i).Item(13)
                        Totals(iTot, 3) += mdsDataSet.Tables(0).Rows(i).Item(14)
                    Next
                Next

                ' Vessel Total
                RowCounter += 1

                '? mdsDataSet.Tables(0).Rows(i).Item(26)
                '"LWQMBY|V"
                '? mdsDataSet.Tables(0).Rows(i).Item(27)
                '"1"
                '? mdsDataSet.Tables(0).Rows(i).Item(28)
                '"Y"

                .SetCellValue(RowCounter, 1, $"Total for Vessel {PrevVessel}:")
                .SetCellValue(RowCounter, 14, Totals(0, 0))
                .SetCellValue(RowCounter, 15, Totals(0, 1))
                .SetCellValue(RowCounter, 16, Totals(0, 2))
                .SetCellValue(RowCounter, 17, Totals(0, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                ' Client Total
                RowCounter += 1
                .SetCellValue(RowCounter, 1, $"Total Payable {PrevClient} {PrevClientName}:")
                .SetCellValue(RowCounter, 14, Totals(1, 0))
                .SetCellValue(RowCounter, 15, Totals(1, 1))
                .SetCellValue(RowCounter, 16, Totals(1, 2))
                .SetCellValue(RowCounter, 17, Totals(1, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                RowCounter += 1
                .SetCellValue(RowCounter, 1, "TOTALS:")
                .SetCellValue(RowCounter, 14, Totals(2, 0))
                .SetCellValue(RowCounter, 15, Totals(2, 1))
                .SetCellValue(RowCounter, 16, Totals(2, 2))
                .SetCellValue(RowCounter, 17, Totals(2, 3))
                .SetCellStyle(RowCounter, 1, RowCounter, ColumnCount, xlBold)

                .AutoFitColumn(1, ColumnCount)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E65_OpsSales(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0

        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Sales")
                .FreezePanes(1, 0)

                Dim xlStyleVoid As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleVoid.Font.FontColor = Color.Gray
                xlStyleVoid.Font.Italic = True

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleGreyOut As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleGreyOut.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleGreyOut.Fill.SetPatternForegroundColor(Color.LightGray)

                Dim xlStyleMandatory As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleMandatory.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleMandatory.Fill.SetPatternForegroundColor(Color.Yellow)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 39, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(17, xlNumStyle)
                .SetColumnStyle(36, 38, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, xlNumStyle)
                .SetColumnStyle(15, xlNumStyle)
                .SetColumnStyle(28, 29, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(9, xlNumStyle)
                .SetColumnStyle(39, xlNumStyle)


                .SetCellValue(1, 1, "Issue Date")
                .SetCellValue(1, 2, "Client Code")
                .SetCellValue(1, 3, "Client Name")
                .SetCellValue(1, 4, "Omit")
                .SetCellValue(1, 5, "Void")
                .SetCellValue(1, 6, "PNR")
                .SetCellValue(1, 7, "Ticket Number")
                .SetCellValue(1, 8, "Passenger")
                .SetCellValue(1, 9, "Pax Count")
                .SetCellValue(1, 10, "Product Type")
                .SetCellValue(1, 11, "Action Type")
                .SetCellValue(1, 12, "Inv Code")
                .SetCellValue(1, 13, "Inv Series")
                .SetCellValue(1, 14, "Inv Number")
                .SetCellValue(1, 15, "Invoice Date")
                .SetCellValue(1, 16, "Vessel")
                .SetCellValue(1, 17, "Net Payable")
                .SetCellValue(1, 18, "Verified")
                .SetCellValue(1, 19, "Remarks")
                .SetCellValue(1, 20, "Transaction Type")
                .SetCellValue(1, 21, "RegNr")
                .SetCellValue(1, 22, "Ticketing Airline")
                .SetCellValue(1, 23, "Routing")
                .SetCellValue(1, 24, "SalesPerson")
                .SetCellValue(1, 25, "Issuing Agent")
                .SetCellValue(1, 26, "Creator Agent")
                .SetCellValue(1, 27, "Reference")
                .SetCellValue(1, 28, "Departure Date")
                .SetCellValue(1, 29, "Arrival Date")
                .SetCellValue(1, 30, "Connected Document")
                .SetCellValue(1, 31, "Pax Remarks")
                .SetCellValue(1, 32, "Doc Status ID")
                .SetCellValue(1, 33, "Cancels Docs")
                .SetCellValue(1, 34, "Sertvices Description")
                .SetCellValue(1, 35, "Client Team")
                .SetCellValue(1, 36, "Sell")
                .SetCellValue(1, 37, "Buy")
                .SetCellValue(1, 38, "Profit")
                .SetCellValue(1, 39, "PaxCount+-")
                .SetCellStyle(1, 1, 1, 39, xlStyleHeader)

                xlINVCount = 1
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    RaiseEvent ProgressCounter(0, mdsDataSet.Tables(0).Rows.Count, i)
                    xlINVCount += 1
                    For j = 0 To 10
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If mdsDataSet.Tables(0).Rows(i).Item(13) <> 0 Then
                        For j = 11 To 14
                            .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                        Next
                    End If
                    For j = 15 To 31
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next
                    If CInt(mdsDataSet.Tables(0).Rows(i).Item(31)) = 43 Then
                        .SetCellValue(xlINVCount, 33, "Cancelled")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 129, xlStyleCancelled)
                    ElseIf CStr(mdsDataSet.Tables(0).Rows(i).Item(32)) <> "" Then
                        .SetCellValue(xlINVCount, 33, $"Cancels {CStr(mdsDataSet.Tables(0).Rows(i).Item(32))}")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 129, xlStyleCancelled)
                    End If
                    For j = 33 To 38
                        .SetCellValue(xlINVCount, j + 1, mdsDataSet.Tables(0).Rows(i).Item(j))
                    Next

                    .SetCellValue(xlINVCount, 34, mdsDataSet.Tables(0).Rows(i).Item(33))
                    .SetCellValue(xlINVCount, 35, mdsDataSet.Tables(0).Rows(i).Item(34))


                    If mdsDataSet.Tables(0).Rows(i).Item(3) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 39, xlStyleOmit)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(4) <> "" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 39, xlStyleVoid)
                    End If
                    If mdsDataSet.Tables(0).Rows(i).Item(10) = "Refund" Then
                        .SetCellStyle(xlINVCount, 1, xlINVCount, 39, xlStyleRefund)
                    End If
                Next

                .AutoFitColumn(1, 39)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Public Sub E66_PurchasesPerAirline(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0
        Dim HeaderRow As Integer = 0
        Dim Totals(18) As Decimal
        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Air Ticket Purchases")

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTitle.Font.Bold = True
                xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleTitle.Fill.SetPatternForegroundColor(Color.LightSteelBlue)
                xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, 1, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(2, 4, xlNumStyle)
                .SetColumnStyle(6, 8, xlNumStyle)
                .SetColumnStyle(10, 13, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0;-#,##0;"
                .SetColumnStyle(5, xlNumStyle)
                .SetColumnStyle(9, xlNumStyle)

                xlINVCount = 2
                .SetCellValue(xlINVCount, 2, $"SALES DATA REPORT - ATHENS {mReport.Date1From:dd/MM/yyyy} - {mReport.Date1To:dd/MM/yyyy}")
                .SetCellStyle(xlINVCount, 2, xlINVCount, 2, xlStyleTitle)
                .MergeWorksheetCells(xlINVCount, 2, xlINVCount, 13)
                xlINVCount += 2
                .SetCellValue(xlINVCount, 1, "Airline")
                .SetCellValue(xlINVCount, 2, "Net CY")
                .SetCellValue(xlINVCount, 3, "Fuel CY")
                .SetCellValue(xlINVCount, 4, "Net+Fuel CY")
                .SetCellValue(xlINVCount, 5, "Cpns CY")
                .SetCellValue(xlINVCount, 6, "Net PY")
                .SetCellValue(xlINVCount, 7, "Fuel PY")
                .SetCellValue(xlINVCount, 8, "Net+Fuel PY")
                .SetCellValue(xlINVCount, 9, "Cpns PY")
                .SetCellValue(xlINVCount, 10, "Index Net")
                .SetCellValue(xlINVCount, 11, "Index Fuel")
                .SetCellValue(xlINVCount, 12, "Index Net+Fuel")
                .SetCellValue(xlINVCount, 13, "Index Cpns")

                .SetCellStyle(xlINVCount, 1, xlINVCount, 13, xlStyleHeader)
                .FreezePanes(xlINVCount, 0)
                HeaderRow = xlINVCount
                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item(0))
                    For ii = 0 To 2
                        For j = 0 To 3
                            If (Not IsDBNull(mdsDataSet.Tables(0).Rows(i).Item(ii * 6 + j + 1))) Then
                                .SetCellValue(xlINVCount, ii * 4 + j + 2, mdsDataSet.Tables(0).Rows(i).Item(ii * 6 + j + 1))
                                Totals(ii * 6 + j + 1) += mdsDataSet.Tables(0).Rows(i).Item(ii * 6 + j + 1)
                            End If
                        Next
                    Next
                Next
                .Sort(HeaderRow + 1, 1, xlINVCount, 13, 6, False)
                .Sort(HeaderRow + 1, 1, xlINVCount, 13, 2, False)
                xlINVCount += 1
                .SetCellValue(xlINVCount, 1, "Totals")
                For ii = 0 To 1
                    For j = 0 To 3
                        .SetCellValue(xlINVCount, ii * 4 + j + 2, Totals(ii * 6 + j + 1))
                    Next
                Next

                If (Totals(7) <> 0) Then .SetCellValue(xlINVCount, 10, Totals(1) / Totals(7) * 100)
                If (Totals(8) <> 0) Then .SetCellValue(xlINVCount, 11, Totals(2) / Totals(8) * 100)
                If (Totals(9) <> 0) Then .SetCellValue(xlINVCount, 12, Totals(3) / Totals(9) * 100)
                If (Totals(10) <> 0) Then .SetCellValue(xlINVCount, 13, Totals(4) / Totals(10) * 100)
                .SetCellStyle(xlINVCount, 1, xlINVCount, 13, xlStyleHeader)
                .SetCellStyle(HeaderRow, 1, xlINVCount, 1, xlStyleHeader)
                .DrawBorder(HeaderRow, 2, xlINVCount, 5, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, Color.Black)
                .DrawBorder(HeaderRow, 6, xlINVCount, 9, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, Color.Black)
                .DrawBorder(HeaderRow, 10, xlINVCount, 13, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick, Color.Black)

                .AutoFitColumn(1, 13)

            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function GetClassLevel(BookingClass As String, ClassRank As String) As Integer
        Dim cc() As String = ClassRank.Split({"|"}, StringSplitOptions.RemoveEmptyEntries)
        Dim ret As Integer = -1
        If cc.GetUpperBound(0) > 0 AndAlso cc(1).IndexOf(BookingClass) >= 0 Then
            ret = 1
        ElseIf cc.GetUpperBound(0) >= 0 Then
            ret = cc(0).IndexOf(BookingClass) + 1
        End If
        Return ret
    End Function
    Public Sub E67_Columbia(ByRef mReport As Reports.ReportsCollection, ByVal FileName As String)
        Dim xlWorkSheet As New SpreadsheetLight.SLDocument

        Dim xlINVCount As Integer = 0
        Dim ColumnCount As Integer = 13
        Dim OldClientCode As String = ""
        Try

            With xlWorkSheet
                .RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "Statement")
                .FreezePanes(1, 0)

                Dim xlStyleTitle As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleTitle.Font.Bold = True
                xlStyleTitle.Font.FontSize = 16

                Dim xlStyleRefund As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleRefund.Font.FontColor = Color.Red

                Dim xlStyleCancelled As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleCancelled.Font.Italic = True

                Dim xlStyleOmit As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleOmit.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleOmit.Fill.SetPatternForegroundColor(Color.SandyBrown)

                Dim xlStyleHeader As SpreadsheetLight.SLStyle = .CreateStyle
                xlStyleHeader.Font.Bold = True
                xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid)
                xlStyleHeader.Fill.SetPatternForegroundColor(Color.Aqua)
                xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)

                Dim xlNumStyle As SpreadsheetLight.SLStyle = .CreateStyle
                xlNumStyle.FormatCode = "@"
                .SetColumnStyle(1, ColumnCount, xlNumStyle)
                xlNumStyle.FormatCode = "#,##0.00;-#,##0.00;"
                .SetColumnStyle(12, xlNumStyle)
                xlNumStyle.FormatCode = "dd/mm/yyyy"
                .SetColumnStyle(1, 3, xlNumStyle)
                .SetColumnStyle(8, xlNumStyle)
                xlINVCount += 1

                .SetCellValue(xlINVCount, 1, "Booking Date")
                .SetCellValue(xlINVCount, 2, "Ticket Date")
                .SetCellValue(xlINVCount, 3, "Invoice Date")
                .SetCellValue(xlINVCount, 4, "Booked By")
                .SetCellValue(xlINVCount, 5, "Department")
                .SetCellValue(xlINVCount, 6, "Vessel")
                .SetCellValue(xlINVCount, 7, "Invoice Number")
                .SetCellValue(xlINVCount, 8, "Departure Date")
                .SetCellValue(xlINVCount, 9, "Destination")
                .SetCellValue(xlINVCount, 10, "Passenger")
                .SetCellValue(xlINVCount, 11, "Rank")
                .SetCellValue(xlINVCount, 12, "Net Payable")
                .SetCellValue(xlINVCount, 13, "Reason For Travel")
                .SetCellStyle(xlINVCount, 1, xlINVCount, ColumnCount, xlStyleHeader)
                Dim routing As String = ""
                Dim TotalPayable As Decimal = 0D

                For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                    If OldClientCode = "" Or OldClientCode <> mdsDataSet.Tables(0).Rows(i).Item("ClientCode") Then
                        If OldClientCode <> "" Then
                            xlINVCount += 1
                            .SetCellValue(xlINVCount, 1, "Total Payable:")
                            .SetCellValue(xlINVCount, 12, TotalPayable)
                            .SetCellStyle(xlINVCount, 1, xlINVCount, ColumnCount, xlStyleHeader)
                            TotalPayable = 0D
                            xlINVCount += 2
                        End If
                        xlINVCount += 1
                        .SetCellValue(xlINVCount, 1, $"{mdsDataSet.Tables(0).Rows(i).Item("ClientCode")}:{mdsDataSet.Tables(0).Rows(i).Item("ClientName")}")
                        .SetCellStyle(xlINVCount, 1, xlINVCount, ColumnCount, xlStyleTitle)
                        .MergeWorksheetCells(xlINVCount, 1, xlINVCount, ColumnCount)
                        OldClientCode = mdsDataSet.Tables(0).Rows(i).Item("ClientCode")
                    End If
                    xlINVCount += 1
                    .SetCellValue(xlINVCount, 1, mdsDataSet.Tables(0).Rows(i).Item("PNRCreationDate"))
                    .SetCellValue(xlINVCount, 2, mdsDataSet.Tables(0).Rows(i).Item("TicketDate"))
                    .SetCellValue(xlINVCount, 3, mdsDataSet.Tables(0).Rows(i).Item("InvoiceDate"))
                    .SetCellValue(xlINVCount, 4, mdsDataSet.Tables(0).Rows(i).Item("BookedBy"))
                    .SetCellValue(xlINVCount, 5, mdsDataSet.Tables(0).Rows(i).Item("Office"))
                    .SetCellValue(xlINVCount, 6, mdsDataSet.Tables(0).Rows(i).Item("Vessel"))
                    .SetCellValue(xlINVCount, 7, mdsDataSet.Tables(0).Rows(i).Item("InvoiceNumber"))
                    .SetCellValue(xlINVCount, 8, mdsDataSet.Tables(0).Rows(i).Item("DepartureDate"))

                    If mdsDataSet.Tables(0).Rows(i).Item("AirportName") + mdsDataSet.Tables(0).Rows(i).Item("CityName") <> "" Then
                        If mdsDataSet.Tables(0).Rows(i).Item("CityName") <> "" Then
                            routing = mdsDataSet.Tables(0).Rows(i).Item("CityName")
                        Else
                            routing = $"{mdsDataSet.Tables(0).Rows(i).Item("AirportName")}/{mdsDataSet.Tables(0).Rows(i).Item("CityName")}"
                        End If
                    Else
                        routing = mdsDataSet.Tables(0).Rows(i).Item("Routing")
                    End If
                    .SetCellValue(xlINVCount, 9, routing)
                    .SetCellValue(xlINVCount, 10, mdsDataSet.Tables(0).Rows(i).Item("Passenger"))
                    .SetCellValue(xlINVCount, 11, mdsDataSet.Tables(0).Rows(i).Item("Rank"))
                    .SetCellValue(xlINVCount, 12, mdsDataSet.Tables(0).Rows(i).Item("NetPayable"))
                    .SetCellValue(xlINVCount, 13, mdsDataSet.Tables(0).Rows(i).Item("ReasonForTravel"))

                    TotalPayable += mdsDataSet.Tables(0).Rows(i).Item("NetPayable")

                Next
                xlINVCount += 1
                .SetCellValue(xlINVCount, 1, "Total Payable:")
                .SetCellValue(xlINVCount, 12, TotalPayable)
                .SetCellStyle(xlINVCount, 1, xlINVCount, ColumnCount, xlStyleHeader)
                .AutoFitColumn(1, ColumnCount)
            End With

            xlWorkSheet.SaveAs(FileName)
            MessageBox.Show("File saved: " & FileName, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function CheckNullZero(ByVal Value As Object) As Boolean

        Try
            If IsDBNull(Value) Then
                Return True
            ElseIf CDec(Value) = 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return True
        End Try

    End Function

End Class
