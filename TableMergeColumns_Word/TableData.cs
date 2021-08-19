using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableMergeColumns_Word
{
    public class TableData
    {
        public int SerialNumber { get; set; }
        public string Measure { get; set; }
        public string ScopeMeasure { get; set; }
        public decimal RecognizedCosts { get; set; }
        public decimal AmountIncentive { get; set; }
        public string MergeDescriptionAdds { get; set; }
    }
}
