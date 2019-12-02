using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VEE2
{
    public class AnalysisDataModel
    {
        public DateTime ReadOutDate = new DateTime();
        public string TransferDate = null;
        public string Obis = null;
        public string Value = null;
        public string ObisFarciDesc = null;
        public string Date = "";

        public AnalysisDataModel()
        {

        }
        public AnalysisDataModel(DateTime readOutDate, string transferDate, string obis, string value, string obisFarciDesc, string date)
        {
            this.ReadOutDate = readOutDate;
            this.TransferDate = transferDate;
            this.Obis = obis;
            this.Value = value;
            this.ObisFarciDesc = obisFarciDesc;
            this.Date = date;
        }
    }
}
