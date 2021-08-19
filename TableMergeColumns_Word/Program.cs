using System;
using System.Collections.Generic;
using System.IO;

namespace TableMergeColumns_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            var AmortizationPlan_WordGenerator = new DocumentGenerator(
                aMeasureMainName: "Ukrep",
                aScopeMeasureMainName: "obseg ukrepa",
                aRecognizedCostsMainName: "priznani stroški",
                aAmountIncentiveMainName: "višina spodbude",
                aTableDatas: new List<TableData>()
                {
                    new TableData()
                    {
                        SerialNumber = 1,
                        Measure = "vgradnja E-TČ po sistemu voda - voda",
                        ScopeMeasure = "(1 kos)",
                        RecognizedCosts = 13868.98m,
                        AmountIncentive = 4000.00m,
                        MergeDescriptionAdds = "Finančna spodbuda je določena glede na z javnim pozivom omejeno višino spodbude na enoto."
                    },
                    new TableData()
                    {
                        SerialNumber = 2,
                        Measure = "vgradnja E-TČ po sistemu voda - voda2",
                        ScopeMeasure = "(2 kos)",
                        RecognizedCosts = 14868.98m,
                        AmountIncentive = 5000.00m,
                        MergeDescriptionAdds = "Finančna spodbuda je določena glede na z javnim pozivom omejeno višino spodbude na enoto."
                    },
                    new TableData()
                    {
                        SerialNumber = 3,
                        Measure = "vgradnja E-TČ po sistemu voda - voda3",
                        ScopeMeasure = "(3 kos)",
                        RecognizedCosts = 15868.98m,
                        AmountIncentive = 6000.00m,
                        MergeDescriptionAdds = "Finančna spodbuda je določena glede na z javnim pozivom omejeno višino spodbude na enoto."
                    }
                },
                aIncentiveSumName: "Spodbuda skupaj",
                aIncentiveSumData: 4000.00m);

            var time1 = DateTime.Now.ToFileTime().ToString();
            File.WriteAllBytes($"C:\\Users\\aleks\\Desktop\\DeloOdDoma\\test\\TableMergeColumnText{time1}.docx", AmortizationPlan_WordGenerator.Generate());
        }
    }
}
