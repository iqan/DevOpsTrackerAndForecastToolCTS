using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProc
{
    class Resource
    {
        public long ProjectId { get; set; }
        public string ProjectName { get; set; }
        public string ResourceName { get; set; }
        public string BillingPeriod { get; set; }
        public int Rate { get; set; }
        public int Leaves { get; set; }
        public int BillingDays { get; set; }
        public int TotalBilling { get; set; }
    }
}
