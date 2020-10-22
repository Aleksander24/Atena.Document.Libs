using System;
using System.Collections.Generic;
using System.Text;

namespace Atena.SupportLibs.DocGenerators.SUB_ListOfRecipient_Word.Models
{
    public class ReceiverData
    {
        public int ZapStevilka { get; set; }
        public string NaslovPrejemnika { get; set; }
        public string PrejemnikSpodbude { get; set; }
        
        public List<Namen> Actions { get; set; }
    }

    public class Namen
    {
        public string NazivNamena { get; set; }
        public string OpisKolicine { get; set; }
        public decimal Velikost { get; set; }
        public string Oznaka { get; set; }
        public decimal VisinaSpodbude { get; set; }
    }
}
