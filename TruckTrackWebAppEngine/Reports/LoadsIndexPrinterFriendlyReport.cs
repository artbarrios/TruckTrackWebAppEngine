using TruckTrackWeb.Models;
using TruckTrackWebAppEngine.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine.Reports
{
    class LoadsIndexPrinterFriendlyReport
    {
        public static void Generate(Report report, string fileSaveDirectory, Application app)
        {
            // generates the report in the specified reportFormat with the
            // specified report.Filename saves it in fileSaveDirectory and always overwrites it
            string saveFilename = Path.Combine(fileSaveDirectory.TrimEnd('\\'), report.Filename.TrimStart('\\')) + "." + report.Extension;
            // gen up the Word objects we need
            Document document = app.Documents.Add();
            // load our styles into the document
            ReportCommon.LoadDocumentStyles(document);

            try
            {
                // build the report document
                // set the document properties
                ReportCommon.SetDocumentDefaultProperties(document, app);
                // add header
                AddDocumentHeader(document);
                // add body
                AddDocumentBody(document);
                // save the document
                document.SaveAs2(saveFilename, report.SaveFormat);
                // display ready message
                AppCommon.Log(report.Name + " ready. Open at: " + AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + report.Extension, AppCommon.GetAppEnginePort()) + " .", EventLogEntryType.Information);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("LoadsIndexViewReport.Generate: " + e.Message, e);
                message += " - Filename = " + saveFilename + "";
                throw new Exception(message);
            }
            finally
            {
                // close and dispose of the writer if it exists
                document.Close(WdSaveOptions.wdDoNotSaveChanges);
            }

        } // Generate

        private static void AddDocumentHeader(Document document)
        {
            // adds the specified part to the document
            // gen up the Word objects we need
            Paragraph paragraph;

            // get a handle to the last paragraph
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Title"]);
            paragraph.Range.Text = "Loads";

            // add trailing blank line
            document.Paragraphs.Add();
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Normal"]);
            paragraph.Range.Text = "";

        } // AddDocumentHeader()

        private static void AddDocumentBody(Document document)
        {
            // adds the specified part to the document
            // gen up the Word objects we need
            Paragraph paragraph;
            Table table;

            // get the data we need to build the report
            List<Load> loadsWebData = new List<Load>();
            loadsWebData = Web_Data.LoadsWebData.GetLoads();

            // add paragraph and get a handle to it
            document.Paragraphs.Add();
            paragraph = document.Paragraphs[document.Paragraphs.Count];
            paragraph.set_Style(document.Styles["Normal"]);

            // add a table and get a handle to it
            document.Tables.Add(paragraph.Range, 1, 3); // 1 X count of properties
            table = document.Tables[document.Tables.Count];
            table.set_Style(document.Styles["Plain Table 2"]);

            // set column widths
            // Example: table.Columns[1].SetWidth(app.InchesToPoints(.75f), WdRulerStyle.wdAdjustSameWidth);
            // set for no in-table page break
            table.Rows[table.Rows.Count].AllowBreakAcrossPages = 0;

            // add column headers
            // Example: table.Rows[table.Rows.Count].Cells[1].Range.Text = "Subject";
            table.Rows[table.Rows.Count].Cells[1].Range.Text = "Name";
            table.Rows[table.Rows.Count].Cells[2].Range.Text = "Truck";
            table.Rows[table.Rows.Count].Cells[3].Range.Text = "Is Time Sensitive";


            // format header row
            table.Rows[table.Rows.Count].HeadingFormat = -1;
            table.Rows[table.Rows.Count].Range.set_Style(document.Styles["TableHeaderRow"]);
            table.Rows[table.Rows.Count].Range.Bold = 1;

            // add table data rows
            foreach (Load load in loadsWebData)
            {
                table.Rows.Add();
                // format data row
                table.Rows[table.Rows.Count].Range.set_Style(document.Styles["TableDataRow"]);
                table.Rows[table.Rows.Count].Range.Bold = 0;
                // Example: table.Rows[table.Rows.Count].Cells[1].Range.Text = object.Name.ToString();
                table.Rows[table.Rows.Count].Cells[1].Range.Text = load.Name.ToString();
                table.Rows[table.Rows.Count].Cells[2].Range.Text = load.TruckId.ToString();
                table.Rows[table.Rows.Count].Cells[3].Range.Text = load.IsTimeSensitive.ToString();

            }

            // add trailing blank line
            paragraph.Range.Text += "";

        } // AddDocumentBody()

    } // class LoadsIndexPrinterFriendlyReport
}

