using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TruckTrackWebAppEngine.Models
{
    class Report
    {
        public string Name { get; set; }
        public string Filename { get; set; }
        public string Extension { get; set; }
        public WdSaveFormat SaveFormat { get; set; }
        public string Url { get; set; }
    }
}

