using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine
{
    class ReportCommon
    {
        public static void LoadDocumentStyles(Document document)
        {
            // adds the needed Styles to the specified document
            // see list of Word Builtin Styles at
            // https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdbuiltinstyle(v=office.15).aspx

            // Normal (built in style that is also the default base style for all newly added styles)
            Style style = document.Styles["Normal"];
            // set the Font
            style.Font.Bold = 0;
            style.Font.Size = 12;
            style.Font.Name = "Calibri";
            // set the ParagraphFormat
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            style.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            style.ParagraphFormat.SpaceAfter = 0;

            // Header
            style = document.Styles["Header"];
            style.Font.Bold = 1;
            style.Font.Size = 32;
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            // TableHeaderRow
            style = document.Styles.Add("TableHeaderRow");
            style.Font.Bold = 1;

            // TableDataRow
            style = document.Styles.Add("TableDataRow");
            style.Font.Bold = 0;

            // CustomStyle
            style = document.Styles.Add("CustomStyle");
            style.Font.Size = 100;

        } // LoadStyles

        public static void SetDocumentDefaultProperties(Document document, Application app)
        {
            // sets the default properties on the specified document 
            // these are the properties that are common to all documents

            // set the document properties
            document.PageSetup.PaperSize = WdPaperSize.wdPaperLetter;
            document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            document.PageSetup.LeftMargin = app.InchesToPoints(0.5f);
            document.PageSetup.RightMargin = app.InchesToPoints(0.5f);
            document.PageSetup.TopMargin = app.InchesToPoints(0.5f);
            document.PageSetup.BottomMargin = app.InchesToPoints(0.5f);

        } // SetDocumentDefaultProperties

    } // class ReportCommon

} // namespace

