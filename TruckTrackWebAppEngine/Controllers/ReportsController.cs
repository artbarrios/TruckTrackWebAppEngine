using TruckTrackWeb.Models;
using TruckTrackWebAppEngine.Models;
using TruckTrackWebAppEngine.Web_Data;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace TruckTrackWebAppEngine
{
    public class ReportsController : ApiController
    {

        // GET /api/reports/SampleReport
        [Route("api/reports/SampleReport")]
        [HttpGet]
        public IHttpActionResult SampleReport()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "SampleReport";
                report.Filename = "SampleReport";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.SampleReport = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // SampleReport()

        // GET /api/reports/VehicleTypesIndexPrinterFriendly
        [Route("api/reports/VehicleTypesIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult VehicleTypesIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "VehicleTypesIndexPrinterFriendly";
                report.Filename = "VehicleTypesIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.VehicleTypesIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // VehicleTypesIndexPrinterFriendly()

        // GET /api/reports/PalletsIndexPrinterFriendly
        [Route("api/reports/PalletsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult PalletsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "PalletsIndexPrinterFriendly";
                report.Filename = "PalletsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.PalletsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // PalletsIndexPrinterFriendly()

        // GET /api/reports/LoadsIndexPrinterFriendly
        [Route("api/reports/LoadsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult LoadsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "LoadsIndexPrinterFriendly";
                report.Filename = "LoadsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.LoadsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // LoadsIndexPrinterFriendly()

        // GET /api/reports/DriversIndexPrinterFriendly
        [Route("api/reports/DriversIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult DriversIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "DriversIndexPrinterFriendly";
                report.Filename = "DriversIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.DriversIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // DriversIndexPrinterFriendly()

        // GET /api/reports/StopEventsIndexPrinterFriendly
        [Route("api/reports/StopEventsIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult StopEventsIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "StopEventsIndexPrinterFriendly";
                report.Filename = "StopEventsIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.StopEventsIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // StopEventsIndexPrinterFriendly()

        // GET /api/reports/TrucksIndexPrinterFriendly
        [Route("api/reports/TrucksIndexPrinterFriendly")]
        [HttpGet]
        public IHttpActionResult TrucksIndexPrinterFriendly()
        {
            try
            {
                // create report object, Url is the public location where it can be viewed with a browser
                Report report = new Report();
                report.Name = "TrucksIndexPrinterFriendly";
                report.Filename = "TrucksIndexPrinterFriendly";
                report.SaveFormat = WdSaveFormat.wdFormatPDF;
                report.Extension = AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat);
                report.Url = AppCommon.BuildUrl(AppCommon.GetAppEngineUrl(), report.Filename + "." + AppCommon.GetExtensionFromWdSaveFormat(report.SaveFormat), AppCommon.GetAppEnginePort());
                // generate the report
                ReportManager.GenerateReport(report);
                // return the report properties
                return Ok(report);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ReportsController.TrucksIndexPrinterFriendly = " + e.Message, e);
                AppCommon.Log(message, EventLogEntryType.Error);
                throw new Exception(message);
            }
        } // TrucksIndexPrinterFriendly()

    }
}

