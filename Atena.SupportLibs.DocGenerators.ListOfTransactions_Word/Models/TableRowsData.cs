using System;
using System.Collections.Generic;
using System.Text;

namespace Atena.SupportLibs.DocGenerators.ListOfTransactions_Word.Models
{
    public class TableRowsData
    {
        public int StNakazila { get; set; }
        public string Razpis { get; set; }
        public string PrejemnikNakazila { get; set; }
        public int DavcnaStevilka { get; set; }
        public string Naslov { get; set; }
        public string StevPogodbe{ get; set; }
        public decimal ZnesekPogodbe { get; set; }
        public decimal Razlika { get; set; }
        public string TRR { get; set; }
        public decimal ZnesekNakazila { get; set; }
    }
}
