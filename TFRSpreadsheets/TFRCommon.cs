using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TFRSpreadsheets
{
    public partial class TFRCommon
    {
        const string NAMETotals = "Totals";
        System.Data.DataSet mdsDataSet;
        string[] MonthNames = { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
        string FileName;
        string mReportTitle;

        ReportsNext.ReportsCollection mReport;

        SpreadsheetLight.SLDocument xlWorkSheet;
        SpreadsheetLight.SLStyle xlStyleDate;
        SpreadsheetLight.SLStyle xlStyleInteger;
        SpreadsheetLight.SLStyle xlStyleDecimal;
        SpreadsheetLight.SLStyle xlStyleText;
        SpreadsheetLight.SLStyle xlStyleBold;
        SpreadsheetLight.SLStyle xlStyleHeader;
        SpreadsheetLight.SLStyle xlStyleTotals;
        SpreadsheetLight.SLStyle xlStyleTotalsYellow;
        public TFRCommon(ReportsNext.ReportsCollection report, System.Data.DataSet pDs, string ReportTitle, string filename)
        {
            mReport = report;
            FileName = filename;
            mReportTitle = ReportTitle;
            mdsDataSet = pDs;
            xlWorkSheet = new SpreadsheetLight.SLDocument();
            PrepareStyles();
        }
        public void ExportToExcel()
        {

        }
        void PrepareStyles()
        {
            xlStyleDate = xlWorkSheet.CreateStyle();
            xlStyleDate.FormatCode = "dd/mm/yyyy";

            xlStyleInteger = xlWorkSheet.CreateStyle();
            xlStyleInteger.FormatCode = "#,##0;-#,##0;";

            xlStyleDecimal = xlWorkSheet.CreateStyle();
            xlStyleDecimal.FormatCode = "#,##0.00;-#,##0.00;";

            xlStyleText = xlWorkSheet.CreateStyle();
            xlStyleText.FormatCode = "@";

            xlStyleBold = xlWorkSheet.CreateStyle();
            xlStyleBold.SetFontBold(true);

            xlStyleHeader = xlWorkSheet.CreateStyle();
            xlStyleHeader.Font.Bold = true;
            xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHeader.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 0, 204, 255));
            xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleTotals = xlWorkSheet.CreateStyle();
            xlStyleTotals.Font.Bold = true;
            xlStyleTotals.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlStyleTotals.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);

            xlStyleTotalsYellow = xlWorkSheet.CreateStyle();
            xlStyleTotalsYellow.Font.Bold = true;
            xlStyleTotalsYellow.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlStyleTotalsYellow.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlStyleTotalsYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleTotalsYellow.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);         
        }
        public string E00_Euronav()
        {

            int xlINVCount = 0;
            int xlCNSCount = 0;

            try
            {
                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, "INV");
                for (int i = 0; i <= 1; i++)
                {
                    if (i == 1) { xlWorkSheet.AddWorksheet("CNS"); }
                    xlWorkSheet.FreezePanes(1, 0);
                    xlWorkSheet.SetColumnStyle(1, 14, xlStyleText);
                    xlWorkSheet.SetColumnStyle(3, xlStyleDate);
                    xlWorkSheet.SetColumnStyle(7, xlStyleDecimal);
                    for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                    {
                        xlWorkSheet.SetCellValue(1, j + 1, mdsDataSet.Tables[0].Columns[j].Caption);
                    }
                }

                for (var i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    if (mdsDataSet.Tables[0].Rows[i][0].ToString().StartsWith("C"))
                    {
                        xlWorkSheet.SelectWorksheet("CNS");
                        xlCNSCount += 1;
                        for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                        {
                            xlWorkSheet.SetCellValue(xlCNSCount + 1, j + 1, Convert.ToString(mdsDataSet.Tables[0].Rows[i][j]));
                        }
                    }
                    else
                    {
                        xlWorkSheet.SelectWorksheet("INV");
                        xlINVCount += 1;
                        for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                        {
                            xlWorkSheet.SetCellValue(xlINVCount + 1, j + 1, Convert.ToString(mdsDataSet.Tables[0].Rows[i][j]));
                        }
                    }
                }
                xlWorkSheet.SelectWorksheet("CNS");
                xlWorkSheet.AutoFitColumn(1, mdsDataSet.Tables[0].Columns.Count);
                xlWorkSheet.SelectWorksheet("INV");
                xlWorkSheet.AutoFitColumn(1, mdsDataSet.Tables[0].Columns.Count);
                xlWorkSheet.SaveAs(FileName);
                return $"File saved: {mReportTitle}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string E02_BSPMonthReportbyairline()
        {
            try
            {
                string mPrevSheetName = "";

                int xlINVCount = 0;
                int xlCNSCount = 0;

                xlWorkSheet.RenameWorksheet(SpreadsheetLight.SLDocument.DefaultFirstSheetName, mReport.BSPMonthDate);
                xlWorkSheet.FreezePanes(1, 0);
                xlWorkSheet.SetColumnStyle(1, 5, xlStyleText);
                xlWorkSheet.SetColumnStyle(6, 12, xlStyleDecimal);
                for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                {
                    xlWorkSheet.SetCellValue(1, j + 1, mdsDataSet.Tables[0].Columns[j].Caption);
                }
                xlWorkSheet.AddWorksheet(NAMETotals);
                xlWorkSheet.FreezePanes(1, 0);
                xlWorkSheet.SetColumnStyle(1, 3, xlStyleText);
                xlWorkSheet.SetColumnStyle(4, 10, xlStyleDecimal);
                for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                {
                    xlWorkSheet.SetCellValue(1, j <= 2 ? j + 1 : j - 1, mdsDataSet.Tables[0].Columns[j].Caption);
                }
                mPrevSheetName = NAMETotals;
                for (var i = 0; i < mdsDataSet.Tables[0].Rows.Count; i++)
                {
                    if (mdsDataSet.Tables[0].Rows[i][3].ToString() == "")
                    {
                        if (mPrevSheetName != NAMETotals)
                        {
                            xlWorkSheet.SelectWorksheet(NAMETotals);
                            mPrevSheetName = NAMETotals;
                        }
                        xlCNSCount++;
                        for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                        {
                            xlWorkSheet.SetCellValue(xlCNSCount + 1, j <= 2 ? j + 1 : j - 1, (mdsDataSet.Tables[0].Rows[i][j] == DBNull.Value) ? "" : mdsDataSet.Tables[0].Rows[i][j].ToString());
                        }
                    }
                    if (mPrevSheetName != mReport.BSPMonthDate)
                    {
                        xlWorkSheet.SelectWorksheet(mReport.BSPMonthDate);
                        mPrevSheetName = mReport.BSPMonthDate;
                    }
                    xlINVCount++;
                    for (var j = 0; j < mdsDataSet.Tables[0].Columns.Count; j++)
                    {
                        if (mdsDataSet.Tables[0].Rows[i][j] != DBNull.Value)
                        {
                            xlWorkSheet.SetCellValue(xlCNSCount + 1, j + 1, mdsDataSet.Tables[0].Rows[i][j].ToString());
                        }

                    }
                    if (mdsDataSet.Tables[0].Rows[i][3].ToString() == "")
                    {
                        xlWorkSheet.SetRowStyle(xlINVCount + 1, xlStyleBold);
                    }
                }
                xlWorkSheet.SelectWorksheet(NAMETotals);
                xlWorkSheet.SetRowStyle(xlCNSCount + 1, xlStyleBold);
                xlWorkSheet.AutoFitColumn(0, mdsDataSet.Tables[0].Columns.Count);
                xlWorkSheet.SelectWorksheet(mReport.BSPMonthDate);
                xlWorkSheet.AutoFitColumn(0, mdsDataSet.Tables[0].Columns.Count);
                xlWorkSheet.SaveAs(FileName);
                return $"File saved: {mReportTitle}";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
