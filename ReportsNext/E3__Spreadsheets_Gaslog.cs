using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsNext
{
    public partial class Spreadsheets
    {
        readonly SpreadsheetLight.SLDocument ThisDocument;
        readonly SpreadsheetLight.SLStyle xlStyleHeader;
        readonly SpreadsheetLight.SLStyle xlStyleTotal;
        readonly SpreadsheetLight.SLStyle xlStyleHotPink;
        readonly SpreadsheetLight.SLStyle xlStyleYellowBold;
        readonly SpreadsheetLight.SLStyle xlStyleCyanBold;
        readonly SpreadsheetLight.SLStyle xlStyleGrayItalic;
        readonly SpreadsheetLight.SLStyle xlStyleRedFont;
        readonly SpreadsheetLight.SLStyle xlTextStyle;
        readonly SpreadsheetLight.SLStyle xlDateStyle;
        readonly SpreadsheetLight.SLStyle xlIntPosStyle;
        readonly SpreadsheetLight.SLStyle xlPctPosStyle;
        readonly SpreadsheetLight.SLStyle xlDecimalStyle;
        public Spreadsheets()
        {
            ThisDocument = new SpreadsheetLight.SLDocument();
            xlStyleHeader = ThisDocument.CreateStyle();
            xlStyleHeader.Font.Bold = true;
            xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHeader.Fill.SetPatternForegroundColor(System.Drawing.Color.Aqua);
            xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleHotPink = ThisDocument.CreateStyle();
            xlStyleHotPink.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHotPink.Fill.SetPatternForegroundColor(System.Drawing.Color.HotPink);

            xlStyleRedFont = ThisDocument.CreateStyle();
            xlStyleRedFont.Font.FontColor = System.Drawing.Color.Red;

            xlStyleYellowBold = ThisDocument.CreateStyle();
            xlStyleYellowBold.Font.Bold = true;
            xlStyleYellowBold.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleYellowBold.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);

            xlStyleGrayItalic = ThisDocument.CreateStyle();
            xlStyleGrayItalic.Font.Italic = true;
            xlStyleGrayItalic.Font.FontColor = System.Drawing.Color.Gray;

            xlStyleCyanBold = ThisDocument.CreateStyle();
            xlStyleCyanBold.Font.Bold = true;
            xlStyleCyanBold.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleCyanBold.Fill.SetPatternForegroundColor(System.Drawing.Color.Cyan);

            xlStyleTotal = ThisDocument.CreateStyle();
            xlStyleTotal.Font.Bold = true;
            xlStyleTotal.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleTotal.Fill.SetPatternForegroundColor(System.Drawing.Color.LightBlue);

            xlTextStyle = ThisDocument.CreateStyle();
            xlTextStyle.FormatCode = "@";

            xlDateStyle = ThisDocument.CreateStyle();
            xlDateStyle.FormatCode = "dd/mm/yyyy";

            xlIntPosStyle = ThisDocument.CreateStyle();
            xlIntPosStyle.FormatCode = "#,##0;;";

            xlPctPosStyle = ThisDocument.CreateStyle();
            xlPctPosStyle.FormatCode = "#,##0.00%;;";

            xlDecimalStyle = ThisDocument.CreateStyle();
            xlDecimalStyle.FormatCode = "#,##0.00;-#,##0.00;0.00";
        }
    }
}
