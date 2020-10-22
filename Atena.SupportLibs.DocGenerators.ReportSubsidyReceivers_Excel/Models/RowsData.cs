using System;
using System.Collections.Generic;
using System.Text;

namespace Atena.SupportLibs.DocGenerators.ReportSubsidyReceivers_Excel.Models
{
    public class RowsData
    {
        public string Prejemnik { get; set; }
        public string NaslovPrejemnika { get; set; }
        public string PostaID { get; set; }
        public int DavcnaStevilka { get; set; }
        public string OpisParametra { get; set; }
        public decimal VisinaPomoci { get; set; }
        public string DatumOdlocbe { get; set; }
    }
}
