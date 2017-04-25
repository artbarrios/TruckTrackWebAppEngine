using TruckTrackWeb.Models;
using TruckTrackWebAppEngine.Models;
using TruckTrackWebAppEngine.Web_Data;
using Microsoft.Office.Interop.Word;
using Microsoft.Owin.Hosting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace TruckTrackWebAppEngine
{
    class AppCommon
    {
        public static RichTextBox mainformRichTextBox;
        public static void StartAppEngine(RichTextBox richTextBox)
        {
            // assign the richTextBox global to point to the one on the main form
            mainformRichTextBox = richTextBox;

            try
            {
                // startup the OWIN webserver on the port specified in App.config
                // http://+ means that any host name is valid (ipaddress, localhost, MachineName etc)
                string url = "http://+:" + AppCommon.GetAppEnginePort().ToString();
                WebApp.Start<Startup>(url);
                Log("Starting TruckTrackWeb App Engine", EventLogEntryType.Information);
                Log("App Engine API URL " + BuildUrl(GetAppEngineUrl(), "", GetAppEnginePort()) + " . ", EventLogEntryType.Information);
                Log("Remote server URL " + AppCommon.GetRemoteWebApiUrl() + " . ", EventLogEntryType.Information);
                Log("Saving files to " + GetFileSaveDirectory() + " . ", EventLogEntryType.Information);
                // purge old files older than the specified number of hours if purge enabled
                int countOfPurgedFiles = 0;
                if (AppCommon.IsPurgeOldFilesEnabled())
                {
                    countOfPurgedFiles = AppCommon.PurgeOldFiles(GetFileSaveDirectory(), AppCommon.GetPurgeAgeHours());
                    if (countOfPurgedFiles > 0)
                    {
                        AppCommon.Log("Purged " + countOfPurgedFiles.ToString() + " files from " + GetFileSaveDirectory() + " .", EventLogEntryType.Information);
                    }
                }
            }
            catch (Exception e)
            {
                string  message = AppendInnerExceptionMessages("Could not start the App Engine - " + e.Message, e);
                Log(message, EventLogEntryType.Error);
                Log("App Engine stopped.", EventLogEntryType.Error);
            }

        } // StartAppEngine()

        public static string GetAppEngineUrl()
        {
            // returns the public url of the App Engine 
            //string value = Properties.Settings.Default.AppEngineUrl;
            try
            {
                return @"http://" + GetPublicIpAddress() + @"/";
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        } // GetAppEngineUrl

        public static string GetPublicIpAddress()
        {
            // returns the public ip address or an empty string if error
            WebClient client = new WebClient();
            string url = @"https://api.ipify.org";
            try
            {
                return client.DownloadString(url);
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Could not discover the public IP address using <" + url + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // GetPublicIpAddress

        public static int GetAppEnginePort()
        {
            // retrieves the value from the App.config file 
            // if missing retuns 9000 as default
            string value = ConfigurationManager.AppSettings["AppEnginePort"];
            try
            {
                return (value != null & value.Length > 0) ? Convert.ToInt32(value) : 9000;
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid AppEnginePort specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // GetAppEnginePort

        public static string GetFileSaveDirectory()
        {
            // retrieves the value from the App.config file 
            // if missing retuns C:\AppEngineFiles as default
            // check if we have a valid Path
            string value = ConfigurationManager.AppSettings["FileSaveDirectory"];
            try
            {
                value = (value != null & value.Length > 0) ? value : @"C:\AppEngineFiles";
                value = Path.GetFullPath(value);
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid FileSaveDirectory specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
            // we have the fileSaveDirectory so return it
            return value;
        } // GetFileSaveDirectory

        public static string GetEventLogName()
        {
            // retrieves the value from the App.config file 
            // if missing retuns AppEngineEventLog as default
            // check if we have a valid Path
            string value = ConfigurationManager.AppSettings["EventLogName"];
            try
            {
                value = (value != null & value.Length > 0) ? value : @"AppEngineEventLog";
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid EventLogName specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
            // we have the fileSaveDirectory so return it
            return value;
        } // GetEventLogName

        public static string GetRemoteWebApiUrl()
        {
            // retrieves the value from the App.config file 
            // if missing retuns empty string as default
            string value = ConfigurationManager.AppSettings["RemoteWebApiUrl"];
            return (value != null & value.Length > 0) ? value : "";
        } // GetRemoteWebApiUrl

        public static int GetPurgeAgeHours()
        {
            // retrieves the value from the App.config file 
            // if missing retuns 0 as default
            string value = ConfigurationManager.AppSettings["PurgeAgeHours"];
            try
            {
                return (value != null & value.Length > 0) ? Convert.ToInt32(value) : 0;
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid PurgeAgeHours specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // GetPurgeAgeHours

        public static bool IsPurgeOldFilesEnabled()
        {
            // returns true if the App.config setting PurgeOldFiles is set to true
            // otherwise returns false
            string value = ConfigurationManager.AppSettings["PurgeOldFiles"];
            try
            {
                bool isPurgeOldFilesEnabled = false;
                if (value != null & value.Length > 0)
                {
                    isPurgeOldFilesEnabled = Convert.ToBoolean(value);
                }
                return isPurgeOldFilesEnabled;
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid PurgeOldFiles value specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // IsPurgeOldFilesEnabled()

        public static bool IsDetailedLogEnabled()
        {
            // returns true if the App.config setting EnableDetailedLogs is set to true
            // otherwise returns false
            string value = ConfigurationManager.AppSettings["EnableDetailedLogs"];
            try
            {
                bool isDetailedLogEnabled = false;
                if (value != null & value.Length > 0)
                {
                    isDetailedLogEnabled = Convert.ToBoolean(value);
                }
                return isDetailedLogEnabled;
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Invalid EnableDetailedLogs value specified <" + value + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // IsDetailedLogEnabled()

        public static string BuildUrl(string host, string path)
        {
            // takes the parameters and returns a valid Url based on them
            // returns empty string if error
            if (!host.ToUpper().StartsWith("HTTP://") && !host.ToUpper().StartsWith("HTTPS://")) { host = "http://" + host; }
            UriBuilder uriBuilder = new UriBuilder();
            // check for valid inputs
            if (host.Length == 0 || path.Length == 0) return "";
            try
            {
                Uri hostUri = new Uri(host);
                uriBuilder.Scheme = hostUri.Scheme;
                uriBuilder.Host = hostUri.Host;
                uriBuilder.Port = hostUri.Port;
                uriBuilder.Path = path;
                return uriBuilder.Uri.ToString();
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("BuildUrl: Could not build URL from these parts <" + host + "> <" + path + ">", e);
                throw new Exception(message);
            }
        } // BuildUrl()

        public static string BuildUrl(string host, string path, int port)
        {
            // takes the parameters and returns a valid Url based on them
            // returns empty string if error
            if (!host.ToUpper().StartsWith("HTTP://") && !host.ToUpper().StartsWith("HTTPS://")) { host = "http://" + host; }
            UriBuilder uriBuilder = new UriBuilder();
            // check for valid inputs
            if (host.Length == 0) return "";
            try
            {
                Uri hostUri = new Uri(host);
                uriBuilder.Scheme = hostUri.Scheme;
                uriBuilder.Host = hostUri.Host;
                uriBuilder.Port = port;
                uriBuilder.Path = path;
                return uriBuilder.Uri.ToString();
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("BuildUrl: Could not build URL from these parts <" + host + " " + path + " " + port.ToString() + "> - " + e.Message, e);
                throw new Exception(message);
            }
        } // BuildUrl()

        public static bool IsUrlValid(string url)
        {
            try
            {
                // checks to see if we can connect to a running web server at url
                // create a webrequest, set timeout to 10 secs max and get the header information
                HttpWebRequest request = HttpWebRequest.Create(url) as HttpWebRequest;
                request.Timeout = 10000;
                request.Method = "HEAD";
                // pull the request and get the response
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                // get the status code of the response
                int statusCode = (int)response.StatusCode;
                // any code between 100 and 400 means the request is good between 500 and 510 is an error
                if (statusCode >= 100 && statusCode < 400)
                {
                    return true;
                }
                else if (statusCode >= 500 && statusCode <= 510)
                {
                    return false;
                }
            }
            catch (WebException ex)
            {
                if (ex.Status == WebExceptionStatus.ProtocolError) //400 errors
                {
                    return false;
                }
                else
                {
                    throw new Exception(ex.Message);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return false;
        }

        public static string GetExtensionFromWdSaveFormat(WdSaveFormat saveFormat)
        {
            // returns the standard WdSaveFormat type based on the reportFormat string
            if (saveFormat == WdSaveFormat.wdFormatDocumentDefault) { return "docx"; };
            if (saveFormat == WdSaveFormat.wdFormatDocument) { return "doc"; };
            if (saveFormat == WdSaveFormat.wdFormatDocument97) { return "doc"; };
            if (saveFormat == WdSaveFormat.wdFormatRTF) { return "rtf"; };
            if (saveFormat == WdSaveFormat.wdFormatPDF) { return "pdf"; };
            // invalid saveFormat so return empty string
            return "";
        }

        public static void Log(string logMessage, EventLogEntryType type)
        {
            // writes message to the specified listbox and the EventLog as an Information type
            int eventId = 0; // type 0 = error, 2 = information
            if (type == EventLogEntryType.Error) { eventId = 0; }
            if (type == EventLogEntryType.Information) { eventId = 2; }
            try
            {
                string prefix = DateTime.Now.ToString("MM/dd/yy HH:MM:ss tt - ");
                if (type == EventLogEntryType.Error) { prefix += "ERROR: "; };
                if (mainformRichTextBox.InvokeRequired)
                {
                    // InvokeRequired is true prior to the instatiation of the UI thread
                    mainformRichTextBox.Invoke((MethodInvoker)(() => mainformRichTextBox.AppendText(prefix + logMessage + Environment.NewLine)));
                }
                else
                {
                    // add message to form
                    mainformRichTextBox.AppendText(prefix + logMessage + Environment.NewLine);
                }
                EventLogManager.WriteEventLog(logMessage, type, eventId);
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("Log: Could not log message - " + e.Message, e);
                throw new Exception(message);
            }
        } // Log()

        public static string AppendInnerExceptionMessages(string message, Exception e)
        {
            // if they exists appends two levels of Exception InnerException.messages to message
            if (e.InnerException != null)
            {
                message += " - " + e.InnerException.Message;
                if (e.InnerException.InnerException != null)
                {
                    message += " - " + e.InnerException.InnerException.Message;
                }
            }
            return message;
        }

        public static int PurgeOldFiles(string directory, int age)
        {
            // deletes the files in directory that are older than 
            // age hours + Now
            int countOfDeletedFiles = 0;
            try
            {
                string[] filenames;
                DateTime oldestAge = DateTime.Now.AddHours(-age);
                //oldestAge = DateTime.Now.AddMinutes(-1);
                DateTime fileAge;
                FileInfo fileInfo;

                // check for valid inputs
                if (Directory.Exists(directory) && age > 0)
                {
                    // get a list of files in directory
                    filenames = Directory.GetFiles(directory);
                    // get the creation date of each filename
                    foreach (string filename in filenames)
                    {
                        fileAge = File.GetLastWriteTime(filename);
                        fileInfo = new FileInfo(filename);
                        // if this is a .tmp file delete it
                        if (Path.GetExtension(filename).ToUpper() == ".TMP")
                        {
                            File.Delete(filename);
                        }
                        else
                        {
                            // check age of the file and if it is not hidden
                            if (fileAge < oldestAge && !fileInfo.Attributes.HasFlag(FileAttributes.Hidden))
                            {
                                // this file is too old so delete it
                                File.Delete(filename);
                                countOfDeletedFiles++;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                string message = AppendInnerExceptionMessages("DeleteOldFiles: Could not delete old files - " + e.Message, e);
                throw new Exception(message);
            }
            return countOfDeletedFiles;

        } // DeleteOldFiles()

        public static string GetMD5Hash(string source)
        {
            // returns the MD5 hash of the specified source string
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] digest = md5.ComputeHash(Encoding.UTF8.GetBytes(source));
            string base64digest = Convert.ToBase64String(digest, 0, digest.Length);
            return base64digest.Substring(0, base64digest.Length - 2);
        }

    } // class AppCommon
}

