using TruckTrackWebAppEngine.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine.Reports
{
    class SampleReport
    {
        public static void Generate(Report report, string fileSaveDirectory, Application app)
        {
            // generates the report in the specified reportFormat with the
            // specified report.Filename saves it in fileSaveDirectory and always overwrites it
            string cleanFilename = report.Filename.TrimStart('\\');
            string cleanPath = fileSaveDirectory.TrimEnd('\\');

            // gen up the Word objects we need
            Document document = app.Documents.Add();
            Paragraph paragraph;

            // load our styles into the document
            ReportCommon.LoadDocumentStyles(document);

            try
            {
                // build the report document

                // add header
                // get a handle to the first paragraph
                paragraph = document.Paragraphs[document.Paragraphs.Count];
                paragraph.set_Style(document.Styles["Header"]);
                paragraph.Range.Text = "Sample Report";
                paragraph.Range.Text += "Header";

                // add body
                // add paragraph and get a handle to the paragraph
                document.Paragraphs.Add();
                paragraph = document.Paragraphs[document.Paragraphs.Count];
                paragraph.set_Style(document.Styles["CustomStyle"]);
                paragraph.Range.Text = "Report Body";

                // save the document
                document.SaveAs2(Path.Combine(cleanPath, cleanFilename) + "." + report.Extension, report.SaveFormat);
            }
            catch (Exception e)
            {
                throw new Exception("SampleReport.Generate: " + e.Message);
            }
            finally
            {
                // close and dispose of the writer if it exists
                document.Close(WdSaveOptions.wdDoNotSaveChanges);
            }

        } // Generate
    }
}

