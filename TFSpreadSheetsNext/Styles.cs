using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TFSpreadSheetsNext
{
    internal class Styles
    {
        internal SpreadsheetLight.SLStyle xlStyleDate { get; }
        internal SpreadsheetLight.SLStyle xlStyleTime { get; }
        internal SpreadsheetLight.SLStyle xlStyleInteger { get; }
        internal SpreadsheetLight.SLStyle xlStyleDecimal { get; }
        internal SpreadsheetLight.SLStyle xlStyleDecimalBlankZero { get; }
        internal SpreadsheetLight.SLStyle xlStyleDecimalYellow { get; }
        internal SpreadsheetLight.SLStyle xlStyleText { get; }
        internal SpreadsheetLight.SLStyle xlStyleTextBold { get; }
        internal SpreadsheetLight.SLStyle xlStyleTextBoldCentre { get; }
        internal SpreadsheetLight.SLStyle xlStyleBold { get; }
        internal SpreadsheetLight.SLStyle xlStyleItalic { get; }
        internal SpreadsheetLight.SLStyle xlStyleHeader { get; }
        internal SpreadsheetLight.SLStyle xlStyleBisqueWithBorder { get; }
        internal SpreadsheetLight.SLStyle xlStyleHeaderNotes { get; }
        internal SpreadsheetLight.SLStyle xlStyleTotals { get; }
        internal SpreadsheetLight.SLStyle xlStyleTotalsYellow { get; }
        internal SpreadsheetLight.SLStyle xlStyleRedFont { get; }
        internal SpreadsheetLight.SLStyle xlStyleRedFontItalic { get; }
        internal SpreadsheetLight.SLStyle xlStyleGrayItalic { get; }
        internal SpreadsheetLight.SLStyle xlStyleGreenYellow { get; }
        internal SpreadsheetLight.SLStyle xlStyleSpringGreen { get; }
        internal SpreadsheetLight.SLStyle xlStylePowderBlue { get; }
        internal SpreadsheetLight.SLStyle xlStyleSkyBlue { get; }
        internal SpreadsheetLight.SLStyle xlStyleSandyBrown { get; }
        internal SpreadsheetLight.SLStyle xlStyleYellowItalic { get; }
        internal SpreadsheetLight.SLStyle xlStyleTitle { get; }
        internal SpreadsheetLight.SLStyle xlStyleGrey { get; }
        internal SpreadsheetLight.SLStyle xlStyleFixed { get; }
        internal SpreadsheetLight.SLStyle xlStyleTitleBold { get; }
        internal SpreadsheetLight.SLStyle xlStyleBisque { get; }
        internal SpreadsheetLight.SLStyle xlStyleYellowBold { get; }
        internal SpreadsheetLight.SLStyle xlStyleLightGreen { get; }
        internal SpreadsheetLight.SLStyle xlCentred { get; }
        internal SpreadsheetLight.SLStyle xlCyanWithBorder { get; }
        internal SpreadsheetLight.SLStyle xlBoldWithBorder { get; }
        /// <summary>
        /// Pattern:Solid
        /// ForegroundColor:LightSteelBlue
        /// </summary>
        internal SpreadsheetLight.SLStyle xlStyleLightSteelBlue { get; }
        internal SpreadsheetLight.SLStyle xlStyleHoneyDew { get; }
        internal SpreadsheetLight.SLStyle xlStyleLemonChiffon { get; }
        internal SpreadsheetLight.SLStyle xlStyleKhaki { get; }
        internal SpreadsheetLight.SLStyle xlStyleLightGray { get; }
        internal SpreadsheetLight.SLStyle xlStyleYellow { get; }
        internal SpreadsheetLight.SLStyle xlStyleLightGrayItalic { get; }

        internal Styles(SpreadsheetLight.SLDocument xlWorkSheet)
        {
            xlStyleDate = xlWorkSheet.CreateStyle();
            xlStyleDate.FormatCode = "dd/mm/yyyy";

            xlStyleTime = xlWorkSheet.CreateStyle();
            xlStyleTime.FormatCode = "HH:mm";

            xlStyleInteger = xlWorkSheet.CreateStyle();
            xlStyleInteger.FormatCode = "#,##0;-#,##0;";

            xlStyleDecimal = xlWorkSheet.CreateStyle();
            xlStyleDecimal.FormatCode = "#,##0.00;-#,##0.00;";

            xlStyleDecimalBlankZero = xlWorkSheet.CreateStyle();
            xlStyleDecimalBlankZero.FormatCode = "#,##0.00;-#,##0.00;;@";


            xlStyleDecimalYellow = xlWorkSheet.CreateStyle();
            xlStyleDecimalYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleDecimalYellow.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);
            xlStyleDecimalYellow.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General);
            xlStyleDecimalYellow.FormatCode = "#,##0.00;-#,##0.00;";

            xlStyleText = xlWorkSheet.CreateStyle();
            xlStyleText.FormatCode = "@";

            xlStyleTextBold = xlWorkSheet.CreateStyle();
            xlStyleTextBold.FormatCode = "@";
            xlStyleTextBold.SetFontBold(true);

            xlStyleTextBoldCentre = xlWorkSheet.CreateStyle();
            xlStyleTextBoldCentre.FormatCode = "@";
            xlStyleTextBoldCentre.SetFontBold(true);
            xlStyleTextBoldCentre.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleBold = xlWorkSheet.CreateStyle();
            xlStyleBold.SetFontBold(true);

            xlBoldWithBorder = xlWorkSheet.CreateStyle();
            xlBoldWithBorder.Font.Bold = true;
            xlBoldWithBorder.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlBoldWithBorder.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);

            xlStyleItalic = xlWorkSheet.CreateStyle();
            xlStyleItalic.Font.Italic = true;

            xlCentred = xlWorkSheet.CreateStyle();
            xlCentred.Alignment.Horizontal = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;

            xlStyleHeader = xlWorkSheet.CreateStyle();
            xlStyleHeader.Font.Bold = true;
            xlStyleHeader.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHeader.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 0, 204, 255));
            xlStyleHeader.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleBisqueWithBorder = xlWorkSheet.CreateStyle();
            xlStyleBisqueWithBorder.Font.Bold = true;
            xlStyleBisqueWithBorder.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleBisqueWithBorder.Fill.SetPatternForegroundColor(System.Drawing.Color.Bisque);
            xlStyleBisqueWithBorder.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
            xlStyleBisqueWithBorder.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlStyleBisqueWithBorder.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);

            xlCyanWithBorder = xlWorkSheet.CreateStyle();
            xlCyanWithBorder.Font.Bold = true;
            xlCyanWithBorder.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlCyanWithBorder.Fill.SetPatternForegroundColor(System.Drawing.Color.Cyan);
            xlCyanWithBorder.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);
            xlCyanWithBorder.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Black);

            xlStyleHeaderNotes = xlWorkSheet.CreateStyle();
            xlStyleHeaderNotes.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHeaderNotes.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 169, 208, 142));
            xlStyleHeaderNotes.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleTitleBold = xlWorkSheet.CreateStyle();
            xlStyleTitleBold.Font.Bold = true;
            xlStyleTitleBold.Font.FontSize = 16;

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

            xlStyleFixed = xlWorkSheet.CreateStyle();
            xlStyleFixed.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleFixed.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 142, 169, 219));
            xlStyleFixed.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleRedFont = xlWorkSheet.CreateStyle();
            xlStyleRedFont.Font.FontColor = System.Drawing.Color.Red;

            xlStyleRedFontItalic = xlWorkSheet.CreateStyle();
            xlStyleRedFontItalic.Font.FontColor = System.Drawing.Color.Red;
            xlStyleRedFontItalic.Font.Italic = true;

            xlStyleGrayItalic = xlWorkSheet.CreateStyle();
            xlStyleGrayItalic.Font.FontColor = System.Drawing.Color.Gray;
            xlStyleGrayItalic.Font.Italic = true;

            xlStyleYellowBold = xlWorkSheet.CreateStyle();
            xlStyleYellowBold.Font.Bold = true;
            xlStyleYellowBold.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleYellowBold.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);

            xlStyleYellowItalic = xlWorkSheet.CreateStyle();
            xlStyleYellowItalic.Font.Italic = true;
            xlStyleYellowItalic.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleYellowItalic.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);

            xlStyleGreenYellow = xlWorkSheet.CreateStyle();
            xlStyleGreenYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleGreenYellow.Fill.SetPatternForegroundColor(System.Drawing.Color.GreenYellow);

            xlStyleSpringGreen = xlWorkSheet.CreateStyle();
            xlStyleSpringGreen.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleSpringGreen.Fill.SetPatternForegroundColor(System.Drawing.Color.SpringGreen);

            xlStylePowderBlue = xlWorkSheet.CreateStyle();
            xlStylePowderBlue.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStylePowderBlue.Fill.SetPatternForegroundColor(System.Drawing.Color.PowderBlue);

            xlStyleSkyBlue = xlWorkSheet.CreateStyle();
            xlStyleSkyBlue.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleSkyBlue.Fill.SetPatternForegroundColor(System.Drawing.Color.SkyBlue);

            xlStyleLightGreen = xlWorkSheet.CreateStyle();
            xlStyleLightGreen.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleLightGreen.Fill.SetPatternForegroundColor(System.Drawing.Color.LightGreen);

            xlStyleSandyBrown = xlWorkSheet.CreateStyle();
            xlStyleSandyBrown.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleSandyBrown.Fill.SetPatternForegroundColor(System.Drawing.Color.SandyBrown);

            xlStyleBisque = xlWorkSheet.CreateStyle();
            xlStyleBisque.Font.Bold = true;
            xlStyleBisque.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleBisque.Fill.SetPatternForegroundColor(System.Drawing.Color.Bisque);
            xlStyleBisque.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left);

            xlStyleGrey = xlWorkSheet.CreateStyle();
            xlStyleGrey.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleGrey.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 191, 191, 191));
            xlStyleGrey.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleTitle = xlWorkSheet.CreateStyle();
            xlStyleTitle.Font.Bold = true;
            xlStyleTitle.Font.FontSize = 15;
            xlStyleTitle.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleTitle.Fill.SetPatternForegroundColor(System.Drawing.Color.FromArgb(255, 191, 191, 191));
            xlStyleTitle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);

            xlStyleLightSteelBlue = xlWorkSheet.CreateStyle();
            xlStyleLightSteelBlue.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleLightSteelBlue.Fill.SetPatternForegroundColor(System.Drawing.Color.LightSteelBlue);

            xlStyleHoneyDew = xlWorkSheet.CreateStyle();
            xlStyleHoneyDew.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleHoneyDew.Fill.SetPatternForegroundColor(System.Drawing.Color.Honeydew);

            xlStyleLemonChiffon = xlWorkSheet.CreateStyle();
            xlStyleLemonChiffon.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleLemonChiffon.Fill.SetPatternForegroundColor(System.Drawing.Color.LemonChiffon);

            xlStyleKhaki = xlWorkSheet.CreateStyle();
            xlStyleKhaki.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleKhaki.Fill.SetPatternForegroundColor(System.Drawing.Color.Khaki);

            xlStyleLightGray = xlWorkSheet.CreateStyle();
            xlStyleLightGray.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleLightGray.Fill.SetPatternForegroundColor(System.Drawing.Color.LightGray);

            xlStyleYellow = xlWorkSheet.CreateStyle();
            xlStyleYellow.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleYellow.Fill.SetPatternForegroundColor(System.Drawing.Color.Yellow);

            xlStyleLightGrayItalic = xlWorkSheet.CreateStyle();
            xlStyleLightGrayItalic.Font.Italic = true;
            xlStyleLightGrayItalic.Fill.SetPatternType(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid);
            xlStyleLightGrayItalic.Fill.SetPatternForegroundColor(System.Drawing.Color.LightGray);
        }
    }
}
