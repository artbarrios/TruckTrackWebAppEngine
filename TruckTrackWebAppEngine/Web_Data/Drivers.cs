using TruckTrackWeb.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine.Web_Data
{
    class DriversWebData
    {
        // global static vars
        private static HttpClient client = new HttpClient();

        // GET: api/DriversData
        public static List<Driver> GetDrivers()
        {
            // return the data or perform an action using the remote webApiUrl
            string webApiPath = "api/DriversData";
            string results = "";
            try
            {
                results = client.GetAsync(AppCommon.BuildUrl(AppCommon.GetRemoteWebApiUrl(), webApiPath)).Result.Content.ReadAsStringAsync().Result;
                return JsonConvert.DeserializeObject<List<Driver>>(results);
            }
            catch (Exception e)
            {
                string message = AppCommon.AppendInnerExceptionMessages("GetDrivers: " + e.Message, e);
                throw new Exception(message);
            }
        } // GetDrivers
    }
}

