using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine
{
    class EventLogManager
    {
        // set the name of the log from App Settings
        private static string EventLogName = AppCommon.GetEventLogName();
        private static EventLog AEEventLog;

        public static void CreateEventLog()
        {
            // creates the event log and if it exists deletes it and then creates it
            try
            {
                //// delete the event log
                //if (EventLog.Exists(EventLogName))
                //{
                //    EventLog.Delete(EventLogName);
                //}
                // create the eventlog if it does not exist
                if (!EventLog.Exists(EventLogName))
                {
                    EventSourceCreationData creationData = new EventSourceCreationData(EventLogName, EventLogName);
                    creationData.MachineName = ".";
                    EventLog.CreateEventSource(creationData);
                    // set the event log to overwrite as needed
                    AEEventLog.ModifyOverflowPolicy(OverflowAction.OverwriteAsNeeded, 0);
                }
                // get a handle to the event log
                AEEventLog = new EventLog();
                AEEventLog.Source = EventLogName;
                AEEventLog.Log = EventLogName;
            }
            catch (Exception e)
            {
                // error writing to our custom event log so write an error to the Application log
                EventLog ApplicationLog = new EventLog();
                ApplicationLog.Source = "AppEngineError";
                ApplicationLog.WriteEntry(Process.GetCurrentProcess().ProcessName + ".CreateEventLog: " + e.Message, EventLogEntryType.Error);
                if (e.Message.ToUpper().Contains("NOT ALLOWED"))
                {
                    ApplicationLog.WriteEntry(Process.GetCurrentProcess().ProcessName + " must be run with Administrator privledges.", EventLogEntryType.Error);
                }
                throw new Exception("CreateEventLog: " + e.Message);
            }

        } // CreateEventLog

        public static void WriteEventLog(string message, EventLogEntryType entryType, int eventId)
        {
            // writes the specified message to the event log
            try
            {
                // check if this is an Information message and if EnableDetailedLogs is true 
                if (entryType == EventLogEntryType.Information && !AppCommon.IsDetailedLogEnabled())
                {
                    // IsDetailedLogEnabled is false so do not write Information entries
                    return;
                }
                // all ok so write to the log
                AEEventLog.WriteEntry(message, entryType, eventId);
            }
            catch (Exception e)
            {
                // do not handle this error - if it occurs we simply not write to the event log 
                throw new Exception("WriteEventLog: " + e.Message);
            }

        } // WriteEventLog
    }
}

