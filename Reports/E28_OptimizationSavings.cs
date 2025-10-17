using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using SpreadsheetLight;

namespace Reports
{
    public class E28_OptimizationSavings
    {
        public void CreateSheet()
        {


            try
            {
                using (var xlWorkSheet = new SpreadsheetLight.SLDocument())
                {
                    var xlStyleRefund = xlWorkSheet.CreateStyle();
                    xlStyleRefund.Font.FontColor = System.Drawing.Color.Red;

                    var xlStyleTotals = xlWorkSheet.CreateStyle();
                    xlStyleTotals.Font.Bold = true;
                    xlStyleTotals.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
                    xlStyleTotals.Fill.SetPatternForegroundColor(System.Drawing.Color.LightGray);

                    var xlStyleHeader = xlWorkSheet.CreateStyle();
                    xlStyleHeader.Font.Bold = true;
                    xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
                    xlStyleHeader.Fill.SetPatternForegroundColor(System.Drawing.Color.Aqua);
                    xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

                    var xlNumStyle=xlWorkSheet.CreateStyle();
                    xlNumStyle.FormatCode = "@";

                }
            }
            catch (Exception ex)
            {
                throw new Exception($"E28_OptimizationSavings.CreateSheet\r\n{ex.Message}");
            }

        }
    }
}
