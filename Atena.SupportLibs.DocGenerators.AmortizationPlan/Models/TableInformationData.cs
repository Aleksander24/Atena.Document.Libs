using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Atena.SupportLibs.DocGenerators.AmortizationPlan.Models
{
    public class TableInformationData
    {
        //_loanValueName; // Vrednost kredita:
        //string _timeOfReturnName; // Doba vračanja:
        //string _moratoriumName; // Moratorij:
        //string _interestRateName; // Obrestna mera:
        //string _typeCalculationName; // Način obračuna:
        //string _firstDateLoanPaidName; // datum prvega obroka:
        public double LoanValue { get; set; }
        public int AgeOfReturnLoan { get; set; }
        public int Moratorium { get; set; }
        public double InterestRate { get; set; }
        public string TypeOfCalculationLoan { get; set; }
        public DateTime FirstDateLoanPaid { get; set; }
    }
}
