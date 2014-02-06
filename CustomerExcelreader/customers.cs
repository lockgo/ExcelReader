using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomerExcelreader
{
    public class customers
    {
        public customers()
        {
            Dates = new List<String>();
            Description = new List<String>();
            moneyIN = new List<double>();
            moneyOUT = new List<double>();
            Balance = new List<double>();
            Remarks = new List<String>();
            Events = new List<string>();
            historyCount = 0;

        }
        /// <summary>
        /// These work just fine.
        /// </summary>
        public String Guest { get; set; }
        public String Contact { get; set; }
        public String ClubNumber { get; set; }//GuestNumber
        public String CasioNumber { get; set; }

        /// <summary>
        /// Lists of lists
        /// </summary>
        public List<String> Dates { get; set; }
        public List<String> Description { get; set; }
        public List<double> moneyIN { get; set; }
        public List<double> moneyOUT { get; set; }
        public List<double> Balance { get; set; }
        public List<String> Remarks { get; set; }
        public List<String> Events { get; set; }


        public int historyCount { get; set; }
    }
}
