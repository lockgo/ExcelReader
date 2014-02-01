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

        }
        public String Guest { get; set; }
        public String Contact { get; set; }
        public String ClubNumber { get; set; }//GuestNumber
        public String CasioNumber { get; set; }

        public List<String> Dates { get; set; }
        public List<String> Description { get; set; }
        public List<String> moneyIN { get; set; }
        public List<String> moneyOUT { get; set; }
        public List<String> Balance { get; set; }
        public List<String> Remarks { get; set; }

    }
}
