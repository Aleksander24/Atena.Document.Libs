using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atena.SupportLibs.DocGenerators.AmortizationPlan.Models
{
    public class TableAmortizationData
    {
        //public int Month { get; set; }
        public double InterestPaid { get; set; } // interestPaid
        public double Annuity { get; set; } // anuiteta
        public double MonthlyPay { get; set; } // monthlyPay
        public double Balance { get; set; } // balance
        public DateTime StartLoanDate { get; set; }
        public double Euribor_OM { get; set; }
    }
}
