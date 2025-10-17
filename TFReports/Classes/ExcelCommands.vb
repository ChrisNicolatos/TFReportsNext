Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelCommands
    Private mdsDataSet As DataSet
    Private mReportTitle As String

    Public Sub New(pDs As DataSet, ByVal ReportTitle As String)

        mReportTitle = ReportTitle
        mdsDataSet = pDs

    End Sub
    Private Sub saveFile(ByRef xlApp As Excel.Application, ByRef xlWorkBook As Excel.Workbook, ByVal FileName As String)

        Try
            xlWorkBook.SaveAs(FileName)
            xlWorkBook.Close()
            xlApp.Quit()
            MessageBox.Show("File saved: " & mReportTitle, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            Throw New Exception("File not saved" & vbCrLf & ex.Message)
        End Try

    End Sub

    Public Sub E00_Euronav(ByVal FileName As String)
        ' Euronav
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet
        Dim xlWorkSheetCNSA As Excel.Worksheet

        Dim xlINVCount As Integer = 0
        Dim xlCNSCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "INVS"

            xlWorkSheetINVS.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetCNSA = xlWorkBook.Sheets.Add
            xlWorkSheetCNSA.Name = "CNSA"

            xlWorkSheetCNSA.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetCNSA.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            Dim darrDataINV(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object
            Dim darrDataCNS(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrDataINV(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
                darrDataCNS(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next


            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                If mdsDataSet.Tables(0).Rows(i).Item(0).ToString.StartsWith("C") Then
                    xlCNSCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        darrDataCNS(xlCNSCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j)
                    Next
                Else
                    xlINVCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        darrDataINV(xlINVCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j)
                    Next
                End If
            Next
            Dim s1DataCellHome = xlWorkSheetINVS.Cells(1, 1)
            Dim s1DataCellEnd = xlWorkSheetINVS.Cells(xlINVCount + 3, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1DataCellHome, s1DataCellEnd).Value2 = darrDataINV

            s1DataCellHome = xlWorkSheetCNSA.Cells(1, 1)
            s1DataCellEnd = xlWorkSheetCNSA.Cells(xlCNSCount + 3, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetCNSA.Range(s1DataCellHome, s1DataCellEnd).Value2 = darrDataCNS

            saveFile(xlApp, xlWorkBook, FileName)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetCNSA = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)
            releaseObject(xlWorkSheetCNSA)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub

    Public Sub E02_BSPMonthReportbyairline(ByVal mBSPMonthDate As String, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet
        Dim xlWorkSheetCNSA As Excel.Worksheet

        Dim xlINVCount As Integer = 0
        Dim xlCNSCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' BSP Month Report by airline
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = mBSPMonthDate

            xlWorkSheetINVS.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetCNSA = xlWorkBook.Sheets.Add
            xlWorkSheetCNSA.Name = "Totals"

            xlWorkSheetCNSA.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            ' prepare data aray which will transfer data from datagrid to excel
            Dim darrINV(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object
            Dim darrCNS(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 3) As Object

            ' copy headers to first row
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrINV(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
                If j <= 2 Then
                    darrCNS(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
                ElseIf j >= 5 Then
                    darrCNS(0, j - 2) = mdsDataSet.Tables(0).Columns(j).Caption
                End If
            Next

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                If mdsDataSet.Tables(0).Rows(i).Item(3).ToString = "" Then
                    xlCNSCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        If j <= 2 Then
                            darrCNS(xlCNSCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j)
                        ElseIf j >= 5 Then
                            darrCNS(xlCNSCount, j - 2) = mdsDataSet.Tables(0).Rows(i).Item(j)
                        End If
                    Next
                End If
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrINV(xlINVCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j)
                Next
            Next

            ' transfer data array to worksheet for detail
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrINV

            ' format data columns for numbers and dates
            s1INV = xlWorkSheetINVS.Cells(2, 4)
            s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, 4)
            xlWorkSheetINVS.Range(s1INV, s2INV).NumberFormat = "dd/mm/yyyy;"
            s1INV = xlWorkSheetINVS.Cells(2, 6)
            s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, 12)
            xlWorkSheetINVS.Range(s1INV, s2INV).NumberFormat = "#,##0.00;-#,##0.00;"

            ' transfer data array to worksheet for totals
            Dim s1CNS = xlWorkSheetCNSA.Cells(1, 1)
            Dim s2CNS = xlWorkSheetCNSA.Cells(xlCNSCount + 1, mdsDataSet.Tables(0).Columns.Count - 2)
            xlWorkSheetCNSA.Range(s1CNS, s2CNS).Value2 = darrCNS
            ' format data columns for numbers
            s1CNS = xlWorkSheetCNSA.Cells(2, 4)
            s2CNS = xlWorkSheetCNSA.Cells(xlCNSCount + 1, 10)
            xlWorkSheetCNSA.Range(s1CNS, s2CNS).NumberFormat = "#,##0.00;-#,##0.00;"

            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetCNSA = xlWorkBook.Sheets("sheet1")
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)
            releaseObject(xlWorkSheetCNSA)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E03_BSPFortnightReportbyticket(ByVal mBSPFortDate As String, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet
        Dim xlWorkSheetCNSA As Excel.Worksheet

        Dim xlINVCount As Integer = 0
        Dim xlCNSCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' BSP Fortnight Report by ticket

            Dim pTotsINV(2, 10) As Decimal ' 0=grand total 1=type total 2=air total
            Dim pINVTots0(0) As Integer
            Dim pINVTots1(0) As Integer
            Dim pINVTots2(0) As Integer
            Dim pAirOldINV As String = ""
            Dim pTypeOldINV As String = ""

            Dim pTotsCNS(2, 10) As Decimal
            Dim pCNSTots0(0) As Integer
            Dim pCNSTots1(0) As Integer
            Dim pCNSTots2(0) As Integer
            Dim pAirOldCNS As String = ""
            Dim pTypeOldCNS As String = ""

            ' initialize arrays which keep track of which rows in excel have totals and must be formatted accordingly at the end
            For ii As Integer = pTotsINV.GetLowerBound(0) To pTotsINV.GetUpperBound(0)
                For jj As Integer = pTotsINV.GetUpperBound(1) To pTotsINV.GetUpperBound(1)
                    pTotsINV(ii, jj) = 0
                    pTotsCNS(ii, jj) = 0
                Next
            Next

            ' prepare excel
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = mBSPFortDate & " I"

            xlWorkSheetINVS.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetCNSA = xlWorkBook.Sheets.Add
            xlWorkSheetCNSA.Name = mBSPFortDate & " D"

            xlWorkSheetCNSA.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            ' prepare data aray which will transfer data from datagrid to excel
            Dim darrINV(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object
            Dim darrCNS(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object

            ' copy headers to first row
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrINV(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
                darrCNS(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next

            ' loop through data grid, row by row
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                ' domestic tickets (D in column 2) go into a separate worksheet
                If mdsDataSet.Tables(0).Rows(i).Item(1).ToString = "D" Then
                    ' check for subtotals triggered by change of airline/ticket type
                    If pAirOldCNS <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Or pTypeOldCNS <> mdsDataSet.Tables(0).Rows(i).Item(2).ToString Then
                        ' if previous ticket type is not empty then this is not the first row
                        If pTypeOldCNS <> "" Then
                            ' add a total row in the data array for type total
                            xlCNSCount += 1
                            darrCNS(xlCNSCount, 0) = pAirOldCNS
                            darrCNS(xlCNSCount, 2) = pTypeOldCNS
                            darrCNS(xlCNSCount, 3) = "TOTAL"
                            ' enter values for type total and accumulate values for airline total
                            ' for columns with values (7-16)
                            For jj As Integer = 7 To 16
                                darrCNS(xlCNSCount, jj) = pTotsCNS(2, jj - 7)
                                pTotsCNS(1, jj - 7) += pTotsCNS(2, jj - 7)
                                pTotsCNS(2, jj - 7) = 0
                            Next
                            ' store row number which has type total
                            ReDim Preserve pCNSTots2(pCNSTots2.GetUpperBound(0) + 1)
                            pCNSTots2(pCNSTots2.GetUpperBound(0)) = xlCNSCount
                            ' add a blank row
                            xlCNSCount += 1
                            ' check for subtotal triggered by change of airline
                            If pAirOldCNS <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Then
                                ' if previous airline is not empty then this is not the first row
                                If pAirOldCNS <> "" Then
                                    ' add a total row in the data array for airline total
                                    xlCNSCount += 1
                                    darrCNS(xlCNSCount, 0) = pAirOldCNS
                                    darrCNS(xlCNSCount, 3) = "TOTAL"
                                    ' enter values for airline total and accumulate values for grand total
                                    ' for columns with values (7-16)
                                    For jj As Integer = 7 To 16
                                        darrCNS(xlCNSCount, jj) = pTotsCNS(1, jj - 7)
                                        pTotsCNS(0, jj - 7) += pTotsCNS(1, jj - 7)
                                        pTotsCNS(1, jj - 7) = 0
                                    Next
                                    ' store row number which has airline total
                                    ReDim Preserve pCNSTots1(pCNSTots1.GetUpperBound(0) + 1)
                                    pCNSTots1(pCNSTots1.GetUpperBound(0)) = xlCNSCount
                                    ' add a blank row
                                    xlCNSCount += 1
                                End If
                            End If
                        End If

                    End If
                    ' add a row in the data array for the data
                    xlCNSCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        darrCNS(xlCNSCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                        ' accumulate ticket values to type total except for columns which are percentages (11 and 13)
                        If j >= 7 And j <= 16 And j <> 11 And j <> 13 Then
                            pTotsCNS(2, j - 7) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        End If
                    Next
                    ' store this row's airline and type for comparison to trigger subtotals
                    pAirOldCNS = mdsDataSet.Tables(0).Rows(i).Item(0).ToString
                    pTypeOldCNS = mdsDataSet.Tables(0).Rows(i).Item(2).ToString
                Else
                    ' non domestic tickets are processed here (I in column 2 - actually not D) go into a spearate worksheet
                    ' the process is identical to the above half of the if statement and should be encapsulated 
                    ' TODO Encapsulate the two different worksheets
                    If pAirOldINV <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Or pTypeOldINV <> mdsDataSet.Tables(0).Rows(i).Item(2).ToString Then
                        If pTypeOldINV <> "" Then
                            xlINVCount += 1
                            darrINV(xlINVCount, 0) = pAirOldINV
                            darrINV(xlINVCount, 2) = pTypeOldINV
                            darrINV(xlINVCount, 3) = "TOTAL"
                            For jj As Integer = 7 To 16
                                darrINV(xlINVCount, jj) = pTotsINV(2, jj - 7)
                                pTotsINV(1, jj - 7) += pTotsINV(2, jj - 7)
                                pTotsINV(2, jj - 7) = 0
                            Next
                            ReDim Preserve pINVTots2(pINVTots2.GetUpperBound(0) + 1)
                            pINVTots2(pINVTots2.GetUpperBound(0)) = xlINVCount
                            xlINVCount += 1
                        End If
                        If pAirOldINV <> mdsDataSet.Tables(0).Rows(i).Item(0).ToString Then
                            If pAirOldINV <> "" Then
                                xlINVCount += 1
                                darrINV(xlINVCount, 0) = pAirOldINV
                                darrINV(xlINVCount, 3) = "TOTAL"
                                For jj As Integer = 7 To 16
                                    darrINV(xlINVCount, jj) = pTotsINV(1, jj - 7)
                                    pTotsINV(0, jj - 7) += pTotsINV(1, jj - 7)
                                    pTotsINV(1, jj - 7) = 0
                                Next
                                ReDim Preserve pINVTots1(pINVTots1.GetUpperBound(0) + 1)
                                pINVTots1(pINVTots1.GetUpperBound(0)) = xlINVCount

                                xlINVCount += 1
                            End If
                        End If
                    End If
                    xlINVCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        darrINV(xlINVCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString()
                        If j >= 7 And j <= 16 And j <> 11 And j <> 13 Then
                            pTotsINV(2, j - 7) += mdsDataSet.Tables(0).Rows(i).Item(j)
                        End If
                    Next
                    pAirOldINV = mdsDataSet.Tables(0).Rows(i).Item(0).ToString
                    pTypeOldINV = mdsDataSet.Tables(0).Rows(i).Item(2).ToString
                End If
            Next

            ' END OF MAIN LOOP WHICH LOOPS THROUGH DATAGRID ROWS

            ' Add type and airline totals for the last tickets in worksheet D
            xlCNSCount += 1
            darrCNS(xlCNSCount, 0) = pAirOldCNS
            darrCNS(xlCNSCount, 2) = pTypeOldCNS
            darrCNS(xlCNSCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrCNS(xlCNSCount, jj) = pTotsCNS(2, jj - 7)
                pTotsCNS(1, jj - 7) += pTotsCNS(2, jj - 7)
                pTotsCNS(2, jj - 7) = 0
            Next
            ReDim Preserve pCNSTots2(pCNSTots2.GetUpperBound(0) + 1)
            pCNSTots2(pCNSTots2.GetUpperBound(0)) = xlCNSCount
            xlCNSCount += 1
            xlCNSCount += 1
            darrCNS(xlCNSCount, 0) = pAirOldCNS
            darrCNS(xlCNSCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrCNS(xlCNSCount, jj) = pTotsCNS(1, jj - 7)
                pTotsCNS(0, jj - 7) += pTotsCNS(1, jj - 7)
                pTotsCNS(1, jj - 7) = 0
            Next
            ReDim Preserve pCNSTots1(pCNSTots1.GetUpperBound(0) + 1)
            pCNSTots1(pCNSTots1.GetUpperBound(0)) = xlCNSCount
            xlCNSCount += 1
            xlCNSCount += 1
            darrCNS(xlCNSCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrCNS(xlCNSCount, jj) = pTotsCNS(0, jj - 7)
            Next
            ReDim Preserve pCNSTots0(pCNSTots0.GetUpperBound(0) + 1)
            pCNSTots0(pCNSTots0.GetUpperBound(0)) = xlCNSCount

            ' Add type and airline totals for the last tickets in worksheet I

            xlINVCount += 1
            darrINV(xlINVCount, 0) = pAirOldINV
            darrINV(xlINVCount, 2) = pTypeOldINV
            darrINV(xlINVCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrINV(xlINVCount, jj) = pTotsINV(2, jj - 7)
                pTotsINV(1, jj - 7) += pTotsINV(2, jj - 7)
                pTotsINV(2, jj - 7) = 0
            Next
            ReDim Preserve pINVTots2(pINVTots2.GetUpperBound(0) + 1)
            pINVTots2(pINVTots2.GetUpperBound(0)) = xlINVCount
            xlINVCount += 1
            xlINVCount += 1
            darrINV(xlINVCount, 0) = pAirOldINV
            darrINV(xlINVCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrINV(xlINVCount, jj) = pTotsINV(1, jj - 7)
                pTotsINV(0, jj - 7) += pTotsINV(1, jj - 7)
                pTotsINV(1, jj - 7) = 0
            Next
            ReDim Preserve pINVTots1(pINVTots1.GetUpperBound(0) + 1)
            pINVTots1(pINVTots1.GetUpperBound(0)) = xlINVCount
            xlINVCount += 1
            xlINVCount += 1
            darrINV(xlINVCount, 3) = "TOTAL"
            For jj As Integer = 7 To 16
                darrINV(xlINVCount, jj) = pTotsINV(0, jj - 7)
            Next
            ReDim Preserve pINVTots0(pINVTots0.GetUpperBound(0) + 1)
            pINVTots0(pINVTots0.GetUpperBound(0)) = xlINVCount

            ' transfer data array to worksheet for tickets type I
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrINV

            ' transfer data array to worksheet for tickets type D
            Dim s1CNS = xlWorkSheetCNSA.Cells(1, 1)
            Dim s2CNS = xlWorkSheetCNSA.Cells(xlCNSCount + 1, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetCNSA.Range(s1CNS, s2CNS).Value2 = darrCNS

            ' Prepare styles for formatting totals rows
            ' 0 for grand total
            Dim pStyle0 As Excel.Style = xlWorkBook.Styles.Add("p0")
            pStyle0.Font.Bold = True
            pStyle0.Interior.Color = My.Settings.GrandTotalColour
            ' 1 for airline total
            Dim pStyle1 As Excel.Style = xlWorkBook.Styles.Add("p1")
            pStyle1.Font.Bold = True
            pStyle1.Interior.Color = My.Settings.AirlineTotalColour
            ' 2 for type total
            Dim pStyle2 As Excel.Style = xlWorkBook.Styles.Add("p2")
            pStyle2.Font.Bold = True
            pStyle2.Interior.Color = My.Settings.TypeTotalColour

            ' loop through arrays of row numbers which have totals and format accordingly in both worksheets
            ' TODO Encapsulate this
            For ii = 1 To pCNSTots0.GetUpperBound(0)
                s1CNS = xlWorkSheetCNSA.Cells(pCNSTots0(ii) + 1, 1)
                s2CNS = xlWorkSheetCNSA.Cells(pCNSTots0(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetCNSA.Range(s1CNS, s2CNS).Style = pStyle0
            Next
            For ii = 1 To pCNSTots1.GetUpperBound(0)
                s1CNS = xlWorkSheetCNSA.Cells(pCNSTots1(ii) + 1, 1)
                s2CNS = xlWorkSheetCNSA.Cells(pCNSTots1(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetCNSA.Range(s1CNS, s2CNS).Style = pStyle1
            Next
            For ii = 1 To pCNSTots2.GetUpperBound(0)
                s1CNS = xlWorkSheetCNSA.Cells(pCNSTots2(ii) + 1, 1)
                s2CNS = xlWorkSheetCNSA.Cells(pCNSTots2(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetCNSA.Range(s1CNS, s2CNS).Style = pStyle2
            Next

            For ii = 1 To pINVTots0.GetUpperBound(0)
                s1INV = xlWorkSheetINVS.Cells(pINVTots0(ii) + 1, 1)
                s2INV = xlWorkSheetINVS.Cells(pINVTots0(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetINVS.Range(s1INV, s2INV).Style = pStyle0
            Next
            For ii = 1 To pINVTots1.GetUpperBound(0)
                s1INV = xlWorkSheetINVS.Cells(pINVTots1(ii) + 1, 1)
                s2INV = xlWorkSheetINVS.Cells(pINVTots1(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetINVS.Range(s1INV, s2INV).Style = pStyle1
            Next
            For ii = 1 To pINVTots2.GetUpperBound(0)
                s1INV = xlWorkSheetINVS.Cells(pINVTots2(ii) + 1, 1)
                s2INV = xlWorkSheetINVS.Cells(pINVTots2(ii) + 1, mdsDataSet.Tables(0).Columns.Count)
                xlWorkSheetINVS.Range(s1INV, s2INV).Style = pStyle2
            Next

            ' format data columns for numbers and dates in both sheets
            ' TODO Encapsulate this
            s1INV = xlWorkSheetINVS.Cells(2, 8)
            s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, 17)
            xlWorkSheetINVS.Range(s1INV, s2INV).NumberFormat = "#,##0.00;-#,##0.00;"
            s1INV = xlWorkSheetINVS.Cells(2, 5)
            s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, 5)
            xlWorkSheetINVS.Range(s1INV, s2INV).NumberFormat = "dd/mm/yyyy;"

            s1CNS = xlWorkSheetCNSA.Cells(2, 8)
            s2CNS = xlWorkSheetCNSA.Cells(xlCNSCount + 1, 17)
            xlWorkSheetCNSA.Range(s1CNS, s2CNS).NumberFormat = "#,##0.00;-#,##0.00;"
            s1CNS = xlWorkSheetCNSA.Cells(2, 5)
            s2CNS = xlWorkSheetCNSA.Cells(xlCNSCount + 1, 5)
            xlWorkSheetCNSA.Range(s1CNS, s2CNS).NumberFormat = "dd/mm/yyyy;"
            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = mBSPFortDate & " I"
            xlWorkSheetCNSA = xlWorkBook.Sheets.Add

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)
            releaseObject(xlWorkSheetCNSA)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E05_ClientTurnover(ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Client Turnover
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Client Turnover"

            xlWorkSheetINVS.Rows(2).Select()
            xlApp.ActiveWindow.FreezePanes = True

            Dim darrINV(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrINV(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrINV(xlINVCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString()
                Next
            Next
            ' transfer data array to worksheet for detail
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrINV

            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E06_ClientSalesPergroup(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal mPrevDateFrom As Date, ByVal mPrevDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Client Sales Per group
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Sales Per Group"



            Dim darrINV(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object

            darrINV(0, 3) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")
            darrINV(0, 4) = Format(mPrevDateFrom, "dd/MM/yyyy") & "-" & Format(mPrevDateTo, "dd/MM/yyyy")
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrINV(0, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next

            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                If mdsDataSet.Tables(0).Rows(i).Item(1).ToString.StartsWith("0") AndAlso (Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(3)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(4)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(5)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(6)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(7)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(8)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(9)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(10)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(11)) Or Not checkNullZero(mdsDataSet.Tables(0).Rows(i).Item(12))) Then
                    xlINVCount += 1
                    For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                        darrINV(xlINVCount, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                    Next
                End If
            Next
            ' transfer data array to worksheet for detail
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 1, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrINV

            xlWorkSheetINVS.Rows(3).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            xlWorkSheetINVS.Columns("I:I").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            xlWorkSheetINVS.Columns("J:J").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"

            xlWorkSheetINVS.Columns("K:K").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("L:L").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"


            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E07_ProfitPerOPSgroup(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal mPrevDateFrom As Date, ByVal mPrevDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Profit Per OPS group
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Profit Per Group"



            Dim darrData(mdsDataSet.Tables(0).Rows.Count + 2, mdsDataSet.Tables(0).Columns.Count)

            darrData(0, 3) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")
            darrData(0, 4) = Format(mPrevDateFrom, "dd/MM/yyyy") & "-" & Format(mPrevDateTo, "dd/MM/yyyy")

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrData(1, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next
            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrData(xlINVCount + 2, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                Next
            Next
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 2, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrData

            xlWorkSheetINVS.Rows(3).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E08_ProfitPerclientgroup(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal mPrevDateFrom As Date, ByVal mPrevDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Profit Per client group
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Profit Per Group"

            Dim darrData(mdsDataSet.Tables(0).Rows.Count + 2, mdsDataSet.Tables(0).Columns.Count)

            darrData(0, 3) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")
            darrData(0, 4) = Format(mPrevDateFrom, "dd/MM/yyyy") & "-" & Format(mPrevDateTo, "dd/MM/yyyy")

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrData(1, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next
            xlINVCount = 0
            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrData(xlINVCount + 2, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                Next
            Next
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 2, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrData

            xlWorkSheetINVS.Rows(3).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("C:C").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E09_ProfitPerOpsGroupExtra(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal mPrevDateFrom As Date, ByVal mPrevDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Profit Per ops group extra
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Profit Per Group"

            Dim darrData(mdsDataSet.Tables(0).Rows.Count + 2, mdsDataSet.Tables(0).Columns.Count)

            darrData(0, 3) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")
            darrData(0, 4) = Format(mPrevDateFrom, "dd/MM/yyyy") & "-" & Format(mPrevDateTo, "dd/MM/yyyy")

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrData(1, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next
            xlINVCount = 0

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrData(xlINVCount + 2, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                Next
            Next
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 2, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrData

            xlWorkSheetINVS.Rows(3).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("C:C").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("I:I").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("J:J").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("K:K").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("L:L").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub
    Public Sub E10_ProfitPerClientGroupExtra(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlINVCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Profit Per client group extra
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Profit Per Group"

            Dim darrData(mdsDataSet.Tables(0).Rows.Count + 2, mdsDataSet.Tables(0).Columns.Count)

            darrData(0, 3) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")

            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrData(1, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next
            xlINVCount = 0

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlINVCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrData(xlINVCount + 2, j) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
                Next
            Next
            Dim s1INV = xlWorkSheetINVS.Cells(1, 1)
            Dim s2INV = xlWorkSheetINVS.Cells(xlINVCount + 2, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1INV, s2INV).Value2 = darrData

            xlWorkSheetINVS.Rows(3).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("C:C").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("I:I").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("J:J").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("K:K").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("L:L").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            saveFile(xlApp, xlWorkBook, FileName)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try
    End Sub

    'Public Sub E11_ProfitPerClientGroupWithPY_ExcelMethod(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)
    '    '
    '    ' THIS IS THE ORIGINAL VERSION THAT TRANSFERRED DATA DIRECTLY TO THE SPREADSHEET
    '    ' Replaced by the version that uses a data array - for speed
    '    ' This version is kept for urgent recovery only
    '    '
    '    Dim xlApp As Excel.Application
    '    Dim xlWorkBook As Excel.Workbook
    '    Dim xlWorkSheetINVS As Excel.Worksheet

    '    Dim xlINVCount As Integer = 0

    '    Dim misValue As Object = System.Reflection.Missing.Value
    '    Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

    '    Try
    '        ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
    '        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

    '        ' Profit Per client group with PY
    '        xlApp = New Excel.Application
    '        xlWorkBook = xlApp.Workbooks.Add(misValue)
    '        xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
    '        xlWorkSheetINVS.Name = "Profit Per Group"

    '        xlWorkSheetINVS.Rows(3).Select()
    '        xlApp.ActiveWindow.FreezePanes = True

    '        xlWorkSheetINVS.Columns("B:B").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "@"

    '        xlWorkSheetINVS.Columns("C:C").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("D:D").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("E:E").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
    '        xlWorkSheetINVS.Columns("F:F").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("G:G").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("H:H").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("I:I").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
    '        xlWorkSheetINVS.Columns("J:J").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("K:K").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("L:L").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
    '        xlWorkSheetINVS.Columns("M:M").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
    '        xlWorkSheetINVS.Columns("N:N").Select()
    '        xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"

    '        xlWorkSheetINVS.Cells(1, 4) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")

    '        For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '            xlWorkSheetINVS.Cells(2, j + 1) = mdsDataSet.Tables(0).Columns(j).Caption
    '        Next
    '        xlINVCount += 1
    '        For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
    '            xlINVCount += 1
    '            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
    '                xlWorkSheetINVS.Cells(xlINVCount + 1, j + 1) = mdsDataSet.Tables(0).Rows(i).Item(j).ToString
    '            Next
    '        Next
    '        saveFile(xlApp, xlWorkBook, FileName)
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    Finally
    '        releaseObject(xlApp)
    '        releaseObject(xlWorkBook)
    '        releaseObject(xlWorkSheetINVS)

    '        ' reset culture which has been set to en-US earlier
    '        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

    '    End Try

    'End Sub
    Public Sub E11_ProfitPerclientgroupwithPY(ByVal mDateFrom As Date, ByVal mDateTo As Date, ByVal FileName As String)

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheetINVS As Excel.Worksheet

        Dim xlDataCount As Integer = 0

        Dim misValue As Object = System.Reflection.Missing.Value
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

        Try
            ' set culture to en-US otherwise Excel can crash on PCs with other language settings i.e. Greek
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            ' Profit Per client group with PY
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")
            xlWorkSheetINVS.Name = "Profit Per Group"

            ' prepare data array which will transfer data from datagrid to excel
            Dim darrData(mdsDataSet.Tables(0).Rows.Count * 2, mdsDataSet.Tables(0).Columns.Count - 1) As Object

            darrData(0, 2) = Format(mDateFrom, "dd/MM/yyyy") & "-" & Format(mDateTo, "dd/MM/yyyy")
            ' copy headers
            For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                darrData(2, j) = mdsDataSet.Tables(0).Columns(j).Caption
            Next

            For i = 0 To mdsDataSet.Tables(0).Rows.Count - 1
                xlDataCount += 1
                For j = 0 To mdsDataSet.Tables(0).Columns.Count - 1
                    darrData(xlDataCount + 2, j) = mdsDataSet.Tables(0).Rows(i).Item(j)
                Next
            Next

            Dim s1DataCellHome = xlWorkSheetINVS.Cells(1, 1)
            Dim s1DataCellEnd = xlWorkSheetINVS.Cells(xlDataCount + 3, mdsDataSet.Tables(0).Columns.Count)
            xlWorkSheetINVS.Range(s1DataCellHome, s1DataCellEnd).Value2 = darrData

            xlWorkSheetINVS.Rows(4).Select()
            xlApp.ActiveWindow.FreezePanes = True

            xlWorkSheetINVS.Columns("B:B").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "@"

            xlWorkSheetINVS.Columns("C:C").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("D:D").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("E:E").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            xlWorkSheetINVS.Columns("F:F").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("G:G").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("H:H").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("I:I").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            xlWorkSheetINVS.Columns("J:J").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("K:K").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("L:L").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"
            xlWorkSheetINVS.Columns("M:M").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0;-#,##0;"
            xlWorkSheetINVS.Columns("N:N").Select()
            xlApp.ActiveWindow.Selection.NumberFormat = "#,##0.00;-#,##0.00;"


            saveFile(xlApp, xlWorkBook, FileName)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheetINVS = xlWorkBook.Sheets("sheet1")

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheetINVS)

            ' reset culture which has been set to en-US earlier
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        End Try

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Function checkNullZero(ByVal Value As Object) As Boolean

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
