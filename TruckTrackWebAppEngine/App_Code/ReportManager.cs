using TruckTrackWebAppEngine.Models;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine
{
    class ReportManager
    {
        // spin up a copy of Word for use during this session
        public static Application app = new Application();
        // specify the directory where all reports are to be located
        private static string fileSaveDirectory = AppCommon.GetFileSaveDirectory();
        // specify the base address for the WebAPI uri
        private static string webApiAddress = AppCommon.GetRemoteWebApiUrl();

        public static void GenerateReport(Report report)
        {
            // generates the specified report in the specified format
            // gives the file the specified filename and stores it in the specified directory

            // check for valid input
            if (report.Name.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Name specified.");
            }
            if (report.Filename.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Filename specified.");
            }
            if (report.Url.Length == 0)
            {
                throw new Exception("GenerateReport: No report.Url specified.");
            }

            // generate the specified report
            AppCommon.Log("Generating report " + report.Name + ".", EventLogEntryType.Information);
            switch (report.Name.ToUpper())
            {
                case "SAMPLEREPORT":
                    Reports.SampleReport.Generate(report, fileSaveDirectory, app);
                    break; // SAMPLEREPORT
                case "VEHICLETYPESINDEXPRINTERFRIENDLY":
                    Reports.VehicleTypesIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "PALLETSINDEXPRINTERFRIENDLY":
                    Reports.PalletsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "LOADSINDEXPRINTERFRIENDLY":
                    Reports.LoadsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "DRIVERSINDEXPRINTERFRIENDLY":
                    Reports.DriversIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "STOPEVENTSINDEXPRINTERFRIENDLY":
                    Reports.StopEventsIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;
                case "TRUCKSINDEXPRINTERFRIENDLY":
                    Reports.TrucksIndexPrinterFriendlyReport.Generate(report, fileSaveDirectory, app);
                    break;

            }

            // purge old files older than the specified number of hours if purge enabled
            int countOfPurgedFiles = 0;
            if (AppCommon.IsPurgeOldFilesEnabled())
            {
                countOfPurgedFiles = AppCommon.PurgeOldFiles(fileSaveDirectory, AppCommon.GetPurgeAgeHours());
                if (countOfPurgedFiles > 0)
                {
                    AppCommon.Log("Purged " + countOfPurgedFiles.ToString() + " files from " + fileSaveDirectory + " .", EventLogEntryType.Information);
                }
            }

        } // GenerateReport

    }
}

