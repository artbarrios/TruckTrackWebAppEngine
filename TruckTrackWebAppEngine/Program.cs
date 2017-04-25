using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TruckTrackWebAppEngine
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                EventLogManager.CreateEventLog();
                EventLogManager.WriteEventLog(@"App Engine " + Process.GetCurrentProcess().ProcessName + @" starting - " + DateTime.Now.ToString(), EventLogEntryType.Information, 1);
                Application.Run(new Form1(Process.GetCurrentProcess().ProcessName));
                EventLogManager.WriteEventLog(@"App Engine " + Process.GetCurrentProcess().ProcessName + @" exiting - " + DateTime.Now.ToString(), EventLogEntryType.Information, 3);
                ReportManager.app.Quit();
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("ERROR: Could not start AppEngine - " + e.Message, e);
                if (message.ToUpper().Contains("NOT ALLOWED"))
                {
                    message += " - Make sure the AppEngine is running with Administrator privledges.";
                }
                MessageBox.Show(message, "App Engine Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

