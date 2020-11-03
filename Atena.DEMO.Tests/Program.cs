using System;
using System.IO;
using System.Collections.Generic;
using Atena.SupportLibs.DocGenerators.ReportSubsidyReceivers_Excel.Models;
using Atena.SupportLibs.DocGenerators.SUB_ListOfRecipient_Word.Models;
using Atena.SupporLibs.DocGenerators.SUB_SPSRequests_Word.Models;
//using Atena.SupportLibs.DocGenerators.ActitvityAnalysis_Word;
using Atena.SupportLibs.DocGenerators.ActitvityAnalysis_Word.GroupsData;
using Atena.SupportLibs.DocGenerators.ListOfTransactions_Word;
using Atena.SupportLibs.DocGenerators.ListOfTransactions_Word.Models;
//using DocumentGenerator = ListOfRemittances_FinishedUnfinished_Word.DocumentGenerator;
using ListOfRemittances_FinishedUnfinished_Word.Models.UnfinishedData;
using ListOfRemittances_FinishedUnfinished_Word.Models.FinishedData;
using Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word;
using Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word.Models;
using System.Drawing;

namespace Atena.DEMO.Tests
{
    class Program
    {
        static void Main(string[] args)
        {


            byte[] faximile = System.IO.File.ReadAllBytes(@"C:\\Users\\Aleksanderv\\source\\repos\\Aleksander24\\Atena.Document.Libs\\Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word\\Images\\ekoskladSignature.png");

            var fundsTransferOrder = new SupportLibs.DocGenerators.FundsTransferOrder_Word.DocumentGenerator(
                aTenderNumber: Environment.NewLine + Environment.NewLine + "\n\n\n\n\n36010-57/2019",
                aInvestment: Environment.NewLine + "NALOŽBA KONČANA",
                aRecipient: "DOMPLAN, d.d.\n" +
                            "Bleiweisova cesta 14\n" +
                            "4000 KRANJ",
                aTransferOrder: "NALOG ZA NAKAZILO SREDSTEV št.: ",
                aTransferOrderBox: "108376/2020",
                aTableTenderDatas: new List<TableTenderData>()
                { 
                    new TableTenderData()
                    {
                        Razpis = "48SUB-SKOB17",
                        PrejemnikSredstev = "DOMPLAN, d.d.",
                        DavcnaStevilka = 66384010,
                        Naslov = "Bleiweisova cesta 14, 4000 KRANJ",
                        StevilkaPogodbe = "36010-57/2019",
                        TRRZaNakazilo = "SI56 0510 0801 0528 081"
                    }
                },
                aDateTransfer: "Datum nakazila: ",
                aAmountTransfer: "Znesek nakazila: ",
                aContractValue: "\n\nPogodbena vrednost:   ",
                aContractValues: 3188.82m,
                aSubtract: "Razlika (pogodba - izplačilo): ",
                aSubtracts: 0.26m,
                aResponsiblePerson1: Environment.NewLine + Environment.NewLine + "\n\n\nVesna Črnilogar\t\t",
                aResponsiblePerson2: "Nevenka Mateja Udovč",
                aPossiblePayment: "Izplačilo za objekt na naslovu ULICA 1. AVGUSTA 9, 11, 4000 KRANJ.",
                aPossibleIncentive: "Spodbuda se izplača v nižjem znesku, ker je račun nižji od ponudbe ob vlogi." + Environment.NewLine,
                aPossibleNotify: Environment.NewLine + "\nObvestiti: DOMPLAN d.d.\n" +
                                 "Bleiweisova cesta 14\n" +
                                 "4000 KRANJ",
                aFaximile: faximile);
                

            var time = DateTime.Now.ToFileTime().ToString();
            File.WriteAllBytes($"C:\\test\\Atena.Documents\\NalogNakaziloSredstev_{time}.doc", fundsTransferOrder.Generate());
        }

        #region ListOfRemittances_FinishedUnfinished
        //static void Main(string[] args)
        //{
        //    var listOfRemittancesFinishedUnfinished_WordGenerator = new DocumentGenerator(
        //        aHead: "SEZNAM NAKAZIL NA DAN",
        //        aDateRemittances: "Datum nakazila",
        //        aUnTenderCode1s: new List<UnTenderCode1>()
        //        {
        //            new UnTenderCode1()
        //            {
        //                ZapStevilka1 = 1,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "UTD1 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode1()
        //            {
        //                ZapStevilka1 = 2,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "UTD1 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode1()
        //            {
        //                ZapStevilka1 = 3,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "UTD1 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        },
        //        aUnTenderCode2s: new List<UnTenderCode2>()
        //        {
        //            new UnTenderCode2()
        //            {
        //                ZapStevilka1 = 4,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "UTD2 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode2()
        //            {
        //                ZapStevilka1 = 5,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "UTD2 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode2()
        //            {
        //                ZapStevilka1 = 6,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "UTD2 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        },
        //        aUnTenderCode3s: new List<UnTenderCode3>()
        //        {
        //            new UnTenderCode3()
        //            {
        //                ZapStevilka1 = 7,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "UTD3 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode3()
        //            {
        //                ZapStevilka1 = 8,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "UTD3 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new UnTenderCode3()
        //            {
        //                ZapStevilka1 = 9,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "UTD3 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        },
        //        aFiTenderCode1s: new List<FiTenderCode1>()
        //        {
        //            new FiTenderCode1
        //            {
        //                ZapStevilka1 = 10,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "FTD1 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode1
        //            {
        //                ZapStevilka1 = 11,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "FTD1 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode1
        //            {
        //                ZapStevilka1 = 12,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "FTD1 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        },
        //        aFiTenderCode2s: new List<FiTenderCode2>()
        //        {
        //            new FiTenderCode2
        //            {
        //                ZapStevilka1 = 13,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "FTD2 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode2
        //            {
        //                ZapStevilka1 = 14,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "FTD2 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode2
        //            {
        //                ZapStevilka1 = 15,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "FTD2 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        },
        //        aFiTenderCode3s: new List<FiTenderCode3>()
        //        {
        //            new FiTenderCode3
        //            {
        //                ZapStevilka1 = 16,
        //                ZapStevilka2 = 1,
        //                Oznaka1 = "FTD3 1",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode3
        //            {
        //                ZapStevilka1 = 17,
        //                ZapStevilka2 = 2,
        //                Oznaka1 = "FTD3 2",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            },
        //            new FiTenderCode3
        //            {
        //                ZapStevilka1 = 18,
        //                ZapStevilka2 = 3,
        //                Oznaka1 = "FTD3 3",
        //                Oznaka2 = "36014-6590/2020",
        //                Prejemnik = "Banka Intesa Sanpaolo d.d."
        //            }
        //        }
        //            );
        //    var time = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"D:\\DeloOdDoma\\test\\SeznamNakazil_DokončanaNedokončana{time}.doc", listOfRemittancesFinishedUnfinished_WordGenerator.Generate()); // popravi v službi
        //}
        #endregion

        #region ListOfTransactions
        //static void Main(string[] args)
        //{
        //    var listOfTransactions_WordGenerator = new DocumentGenerator(
        //        aDate: $"{Environment.NewLine + Environment.NewLine + Environment.NewLine}Spisek nakazil na dan:",
        //        aTableRowsDatas: new List<TableRowsData>()
        //        {
        //            new TableRowsData()
        //            {
        //                StNakazila = 1,
        //                Razpis = "108939/2020 76FS-PO19",
        //                PrejemnikNakazila = "AGENCIJA OSKAR, d.o.o.",
        //                DavcnaStevilka = 65099343,
        //                Naslov = "Zasavska cesta 45D 4000 KRANJ",
        //                StevPogodbe = "36026-374/2019",
        //                ZnesekPogodbe = 14163.67m,
        //                Razlika = 40900.00m,
        //                TRR = "SI56 300000000450019",
        //                ZnesekNakazila = 14163.67m
        //            },
        //            new TableRowsData()
        //            {
        //                StNakazila = 2,
        //                Razpis = "101234/2020 21FS-PO21",
        //                PrejemnikNakazila = "ALEKSANDER PERŠIČ",
        //                DavcnaStevilka = 12345678,
        //                Naslov = "Zaloška cesta 4 1000 LJUBLJANA",
        //                StevPogodbe = "12345-374/2020",
        //                ZnesekPogodbe = 35163.67m,
        //                Razlika = 0.00m,
        //                TRR = "SI56 300000000674398",
        //                ZnesekNakazila = 82163.67m
        //            }
        //        },
        //        aSumTransactions: "Vsota nakazil",


        //        aResponsiblePerson: $"{Environment.NewLine} Vesna Črnilogar",
        //        aResponsiblePerson2: "Nevenka Mateja Udovč");

        //    var timeA = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"D:\\DeloOdDoma\\test\\SeznamNakazil{timeA}.doc", listOfTransactions_WordGenerator.Generate());

        //}
        #endregion

        #region ActivityAnalysis
        //static void Main(string[] args)
        //{
        //    var activityAnalysis_WordGenerator = new SupportLibs.DocGenerators.ActitvityAnalysis_Word.DocumentGenerator(
        //        aHead: "ANALIZA AKTIVNOSTI EKO SKLADA\n",
        //        aHeadObjects12: "OBJEKTI 1+2",
        //        aHeadObjectsVec: "OBJEKTI VEČ.",
        //        aHeadObjectsVis: "OBJEKTI VIS",
        //        aHeadObjectsLs: "OBJEKTI LS",
        //        aHeadObjectsEvpol: "OBJEKTI EVPOL",
        //        aHeadObjectsSamoO: "OBJEKTI SamoO",
        //        aHeadObjectsEnPr: "OBJEKTI EnPr",
        //        aHeadVehiclesFO: "VOZILA FO",
        //        aHeadVehiclesPO: "VOZILA PO",
        //        aHeadVehiclesMunicipality: "VOZILA Občine JP",
        //        aObjects12s: new List<Objects_1_2>()
        //        {
        //            new Objects_1_2()
        //            {
        //                Leto = 2011,
        //                OdobrenoUredba = 9984597M,
        //                NakazanoUredba = 5095311M,
        //                OdobrenoSPS = 0,
        //                NakazanoSPS = 0,
        //                Vlog = 6790m,
        //                Nalozb = 7577m
        //            },
        //            new Objects_1_2()
        //            {
        //                Leto = 2012,
        //                OdobrenoUredba = 22335850M,
        //                NakazanoUredba = 17986471M,
        //                OdobrenoSPS = 0,
        //                NakazanoSPS = 0,
        //                Vlog = 14543m,
        //                Nalozb = 16471m
        //            }
        //        },
        //        aObjectsVecs: new List<ObjectsVec>()
        //        {
        //            new ObjectsVec()
        //            {
        //                Leto1 = 2011,
        //                OdobrenoUredba1 = 3458878m,
        //                NakazanoUredba1 = 1332903m,
        //                OdobrenoSPS1 = 0m,
        //                NakazanoSPS1 = 0m,
        //                Vlog1 = 396,
        //                Nalozb1 = 467
        //            },
        //            new ObjectsVec()
        //            {
        //                Leto1 = 2012,
        //                OdobrenoUredba1 = 3458878m,
        //                NakazanoUredba1 = 1332903m,
        //                OdobrenoSPS1 = 0m,
        //                NakazanoSPS1 = 0m,
        //                Vlog1 = 396,
        //                Nalozb1 = 467
        //            }
        //        },
        //        aObjectsViss: new List<ObjectsVis>()
        //        {
        //            new ObjectsVis
        //            {
        //                Leto2 = 2011,
        //                OdobrenoUredba2 = 12345,
        //                NakazanoUredba2 = 12345,
        //                OdobrenoSPS2 = 0,
        //                NakazanoSPS2 = 0,
        //                Vlog2 = 300,
        //                Nalozb2 = 200
        //            },
        //            new ObjectsVis
        //            {
        //                Leto2 = 2012,
        //                OdobrenoUredba2 = 12345,
        //                NakazanoUredba2 = 12345,
        //                OdobrenoSPS2 = 0,
        //                NakazanoSPS2 = 0,
        //                Vlog2 = 300,
        //                Nalozb2 = 200
        //            }
        //        },
        //        aObjectsLss: new List<ObjectsLs>()
        //        {
        //            new ObjectsLs
        //            {
        //                Leto3 = 2011,
        //                OdobrenoUredba3 = 12345,
        //                NakazanoUredba3 = 12345,
        //                OdobrenoSPS3 = 0,
        //                NakazanoSPS3 = 0,
        //                Vlog3 = 200,
        //                Nalozb3 = 200
        //            },
        //            new ObjectsLs
        //            {
        //                Leto3 = 2012,
        //                OdobrenoUredba3 = 12345,
        //                NakazanoUredba3 = 12345,
        //                OdobrenoSPS3 = 0,
        //                NakazanoSPS3 = 0,
        //                Vlog3 = 200,
        //                Nalozb3 = 200
        //            }
        //        },
        //        aObjectsEvpols: new List<ObjectsEvpol>()
        //        {
        //            new ObjectsEvpol
        //            {
        //                Leto4 = 2011,
        //                OdobrenoUredba4 = 12345,
        //                NakazanoUredba4 = 12345,
        //                OdobrenoSPS4 = 0,
        //                NakazanoSPS4 = 0,
        //                Vlog4 = 140,
        //                Nalozb4 = 120
        //            },
        //            new ObjectsEvpol
        //            {
        //                Leto4 = 2012,
        //                OdobrenoUredba4 = 12345,
        //                NakazanoUredba4 = 12345,
        //                OdobrenoSPS4 = 0,
        //                NakazanoSPS4 = 0,
        //                Vlog4 = 140,
        //                Nalozb4 = 120
        //            }
        //        },
        //        aObjectsSamoOs: new List<ObjectsSamoO>()
        //        {
        //            new ObjectsSamoO
        //            {
        //                Leto5 = 2011,
        //                OdobrenoUredba5 = 21345,
        //                NakazanoUredba5 = 21345,
        //                OdobrenoSPS5 = 0,
        //                NakazanoSPS5 = 0,
        //                Vlog5 = 120,
        //                Nalozb5 = 300
        //            },
        //            new ObjectsSamoO
        //            {
        //                Leto5 = 2012,
        //                OdobrenoUredba5 = 21345,
        //                NakazanoUredba5 = 21345,
        //                OdobrenoSPS5 = 0,
        //                NakazanoSPS5 = 0,
        //                Vlog5 = 120,
        //                Nalozb5 = 300
        //            }
        //        },
        //        aObjectsEnPrs: new List<ObjectsEnPr>()
        //        {
        //            new ObjectsEnPr
        //            {
        //                Leto6 = 2011,
        //                OdobrenoUredba6 = 4567,
        //                NakazanoUredba6 = 3456,
        //                OdobrenoSPS6 = 0,
        //                NakazanoSPS6 = 0,
        //                Vlog6 = 120,
        //                Nalozb6 = 939
        //            },
        //            new ObjectsEnPr
        //            {
        //                Leto6 = 2012,
        //                OdobrenoUredba6 = 4567,
        //                NakazanoUredba6 = 3456,
        //                OdobrenoSPS6 = 0,
        //                NakazanoSPS6 = 0,
        //                Vlog6 = 120,
        //                Nalozb6 = 939
        //            }
        //        },
        //        aVehiclesFOs: new List<VehiclesFO>()
        //        {
        //            new VehiclesFO
        //            {
        //                Leto7 = 2011,
        //                OdobrenoUredba7 = 4567,
        //                NakazanoUredba7 = 3456,
        //                OdobrenoSPS7 = 0,
        //                NakazanoSPS7 = 0,
        //                Vlog7 = 120,
        //                Nalozb7 = 939
        //            },
        //            new VehiclesFO
        //            {
        //                Leto7 = 2012,
        //                OdobrenoUredba7 = 4567,
        //                NakazanoUredba7 = 3456,
        //                OdobrenoSPS7 = 0,
        //                NakazanoSPS7 = 0,
        //                Vlog7 = 120,
        //                Nalozb7 = 939
        //            }
        //        },
        //        aVehiclesPOs: new List<VehiclesPO>()
        //        {
        //            new VehiclesPO
        //            {
        //                Leto8 = 2011,
        //                OdobrenoUredba8 = 4567,
        //                NakazanoUredba8 = 3456,
        //                OdobrenoSPS8 = 0,
        //                NakazanoSPS8 = 0,
        //                Vlog8 = 120,
        //                Nalozb8 = 939
        //            },
        //            new VehiclesPO
        //            {
        //                Leto8 = 2012,
        //                OdobrenoUredba8 = 4567,
        //                NakazanoUredba8 = 3456,
        //                OdobrenoSPS8 = 0,
        //                NakazanoSPS8 = 0,
        //                Vlog8 = 120,
        //                Nalozb8 = 939
        //            }
        //        },
        //        aVehiclesMunicipalityJPs: new List<VehiclesMunicipalityJP>()
        //        {
        //            new VehiclesMunicipalityJP
        //            {
        //                Leto9 = 2011,
        //                OdobrenoUredba9 = 4567,
        //                NakazanoUredba9 = 3456,
        //                OdobrenoSPS9 = 0,
        //                NakazanoSPS9 = 0,
        //                Vlog9 = 120,
        //                Nalozb9 = 939
        //            },
        //            new VehiclesMunicipalityJP
        //            {
        //                Leto9 = 2012,
        //                OdobrenoUredba9 = 4567,
        //                NakazanoUredba9 = 3456,
        //                OdobrenoSPS9 = 0,
        //                NakazanoSPS9 = 0,
        //                Vlog9 = 120,
        //                Nalozb9 = 939
        //            }
        //        },
        //        aSummAllText: "Vse skupaj:");
        //    var timeA = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"C:\\test\\Atena.Documents\\AnalizaAktivnosti_EKOSKLADA_{timeA}.doc", activityAnalysis_WordGenerator.Generate());
        //}
        #endregion

        #region ReportSubsidyReceiver
        //static void Main(string[] args)
        //{
        //    var reportSubsidyReceivers_ExcelGenerator = new SupportLibs.DocGenerators.ReportSubsidyReceivers_Excel.DocumentGenerator(
        //    aReceiver: "Prejemnik",
        //    aAddressReceiver: "NaslovPrejemnika",
        //    aMailID: "PostaID",
        //    aTaxNumber: "DavcnaStevilka",
        //    aParameterDesc: "OpisParametra",
        //    aAmountHelp: "VisinaPomoci_0",
        //    aDateDesicion: "DatumOdlocbe",
        //    aRowDatas: new List<RowsData>()
        //    {
        //        new RowsData()
        //        {
        //            Prejemnik = "P 1",
        //            NaslovPrejemnika = "NP 1",
        //            PostaID = "PID 1",
        //            DavcnaStevilka = 12345678,
        //            OpisParametra = "OP 1",
        //            VisinaPomoci = 11.1m,
        //            DatumOdlocbe = "1.1.2011"
        //        },
        //        new RowsData()
        //        {
        //            Prejemnik = "P 2",
        //            NaslovPrejemnika = "NP 2",
        //            PostaID = "PID 2",
        //            DavcnaStevilka = 12345678,
        //            OpisParametra = "OP 2",
        //            VisinaPomoci = 22.2m,
        //            DatumOdlocbe = "2.2.2022"
        //        },
        //        new RowsData()
        //        {
        //            Prejemnik = "P 3",
        //            NaslovPrejemnika = "NP 3",
        //            PostaID = "PID 3",
        //            DavcnaStevilka = 12345678,
        //            OpisParametra = "OP 3",
        //            VisinaPomoci = 33.3m,
        //            DatumOdlocbe = "3.3.2033"
        //        }
        //    });
        //    var time = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"C:\\test\\Atena.Documents\\ATENA_PoročiloAKT_nepovratne-Test_{time}.xlsx", reportSubsidyReceivers_ExcelGenerator.Generate());
        //}
        #endregion

        #region SUB-SPSRequest_Word
        //static void Main(string[] args)
        //{
        //var SPSRequest_WordGenerator = new SupporLibs.DocGenerators.SUB_SPSRequests_Word.DocumentGenerator(
        //    aSender: "EKO SKLAD,\n" +
        //    "SLOVENSKI OKOLJSKI JAVNI SKLAD\n" +
        //    "BLEIWEISOV CESTA 30\n" +
        //    "Davčna številka: 10677798\n\n",

        //    aReceiver: "REPUBLIKA SLOVENIJA\n" +
        //                "MINISTRISTVO ZA OKOLJE IN PROSTOR\n" +
        //                "DUNAJSKA CESTA 47\n" +
        //                "1000 LJUBLJANA\n" +
        //                "Davčna številka: 31162991\n",

        //    aTransferRequest: "ZAHTEVEK ZA NAKAZILO številka: ",

        //    aTransferRequestCont: "60-SUB/2016",

        //    aDate: "\t\t\tV Ljubljani, dne:",

        //    aPublicTenderText: "\nna podlagi 6. člena pogodbe 2550-16-31100\n" +
        //        "Javni poziv: 37SUB-OB16\n\n\n",

        //    aProgramFunds: "NAKAZILO NEPOVRATNIH SREDSTEV NA TRR: EKO SKLAD, j.s. - PROGRAMSKA SREDSTVA štev: SI56 0110 0695 0960 378\n",

        //    aRowDatas: new List<MainTableRowsData>()
        //    {
        //        new MainTableRowsData()
        //        {
        //            ZapStevilka= 1,
        //            RegularStevilka ="4718",
        //            StevilkaPogodbe = "36014-8158/2017",
        //            Prejemnik = "Prejemnik 92545",
        //            Naslov = "naslov 92545",
        //            Posta = "1000 Ljubljana",
        //            DavcnaStevilka = 12345678,
        //            VrednostVEUR = 600.00m
        //        } ,
        //        new MainTableRowsData() {
        //            ZapStevilka= 2,
        //            RegularStevilka ="4719",
        //            StevilkaPogodbe = "36014-8648/2017",
        //            Prejemnik = "Prejemnik 92551",
        //            Naslov = "naslov 92551",
        //            Posta = "1000 Ljubljana",
        //            DavcnaStevilka = 12345678,
        //            VrednostVEUR = 1882.17M
        //        } ,
        //        new MainTableRowsData() {
        //            ZapStevilka= 3,
        //            RegularStevilka ="4719",
        //            StevilkaPogodbe = "36014-8648/2017",
        //            Prejemnik = "Prejemnik 92551",
        //            Naslov = "naslov 92551",
        //            Posta = "1000 Ljubljana",
        //            DavcnaStevilka = 12345678,
        //            VrednostVEUR = 1882.17M
        //        } ,
        //        new MainTableRowsData() {
        //            ZapStevilka= 4,
        //            RegularStevilka ="4719",
        //            StevilkaPogodbe = "36014-8648/2017",
        //            Prejemnik = "Prejemnik 92551",
        //            Naslov = "naslov 92551",
        //            Posta = "1000 Ljubljana",
        //            DavcnaStevilka = 12345678,
        //            VrednostVEUR = 1882.17M
        //        } ,
        //        new MainTableRowsData() {
        //            ZapStevilka= 5,
        //            RegularStevilka ="4719",
        //            StevilkaPogodbe = "36014-8648/2017",
        //            Prejemnik = "Prejemnik 92551",
        //            Naslov = "naslov 92551",
        //            Posta = "1000 Ljubljana",
        //            DavcnaStevilka = 12345678,
        //            VrednostVEUR = 1882.17M
        //        } ,
        //    },

        //    aSPSRecapitulations: new List<SPSRecapitulationData>()
        //    {
        //        new SPSRecapitulationData
        //        {
        //            SPSProjectName = "2550-17-0021 Ogrevalne naprave (Kurilne+TČ)",
        //            SPSProjectSum = 6602.66M
        //        },
        //        new SPSRecapitulationData
        //        {
        //            SPSProjectName = "2550-17-0022 Ostali ukrepi na stavbah",
        //            SPSProjectSum = 16769.54M
        //        }
        //    },

        //aPrepared: "\nPripravil: mag. Igor Čehovin",
        //aResponsiblePerson: "Odgovorna oseba: mag. Vesna Črnilogar\n",
        //aAttachments: "\n\nPriloge:" +
        //        "\n - pogodbe\n"
        //    );
        //    var time1 = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"C:\\test\\Atena.Documents\\SUB-SPS_Request_{time1}.docx", SPSRequest_WordGenerator.Generate());
        //}
        #endregion

        #region SUB-ListOfRecipient_Word
        //static void Main(string[] args)
        //{
        //var listOfRecipient_WordGenerator = new Atena.SupportLibs.DocGenerators.SUB_ListOfRecipient_Word.DocumentGenerator(
        //    aTextFinancialIncentive: "Seznam prejemnikov nepovratnih finančnih spodbud, ki ga Eko sklad j.s. objavlja na podlagi " +
        //    "316. člena Energetskega zakona EZ1 (Ur.l. RS, št. 17/14) in 3. točke prvega odstavka 10. člena Uredbe o posredovanju " +
        //    "in ponovni uporabi informacij javnega značaja (Ur. l. RS, št. 24/16)\n",

        //    aTextPayouts: "Izplačila nepovratnih finančnih spodbud v letu 2017",

        //    aRowDatas: new List<ReceiverData>()
        //    {
        //        new ReceiverData()
        //        {
        //            ZapStevilka = 1,
        //            PrejemnikSpodbude = "FO 1",
        //            NaslovPrejemnika = "NP 1",

        //            Actions = new List<Namen>()
        //            {
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 1 - Namen 1",
        //                    OpisKolicine = "FO 1 - OK 1",
        //                    Velikost = 1.0M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 10.00M
        //                }
        //            }
        //        },
        //        new ReceiverData()
        //        {
        //            ZapStevilka = 2,
        //            PrejemnikSpodbude = "FO 2",
        //            NaslovPrejemnika = "NP 2",
        //            Actions = new List<Namen>()
        //            {
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 2 - Namen 1",
        //                    OpisKolicine = "FO 2 - OK 1",
        //                    Velikost = 2.0M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 20.00M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 2 - Namen 2",
        //                    OpisKolicine = "FO 2 - OK 2",
        //                    Velikost = 2.1M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 20.10M
        //                }
        //            }
        //        },
        //        new ReceiverData()
        //        {
        //            ZapStevilka = 3,
        //            PrejemnikSpodbude = "FO 3",
        //            NaslovPrejemnika = "NP 3",
        //            Actions = new List<Namen>()
        //            {
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 3 - Namen 1",
        //                    OpisKolicine = "FO 3 - OK 1",
        //                    Velikost = 3.0M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 30.00M
        //                }
        //            }
        //        },
        //        new ReceiverData()
        //        {
        //            ZapStevilka = 4,
        //            PrejemnikSpodbude = "FO 4",
        //            NaslovPrejemnika = "NP 4",
        //            Actions = new List<Namen>()
        //            {
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 4 - Namen 1",
        //                    OpisKolicine = "FO 4 - OK 1",
        //                    Velikost = 4.0M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 40.00M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 4 - Namen 2",
        //                    OpisKolicine = "FO 4 - OK 2",
        //                    Velikost = 4.1M,
        //                    Oznaka ="kom",
        //                    VisinaSpodbude = 40.10M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 4 - Namen 3",
        //                    OpisKolicine = "FO 4 - OK 3",
        //                    Velikost = 4.2M,
        //                    Oznaka ="kW",
        //                    VisinaSpodbude = 40.20M
        //                }
        //            }
        //        },
        //        new ReceiverData()
        //        {
        //            ZapStevilka = 5,
        //            PrejemnikSpodbude = "FO 5",
        //            NaslovPrejemnika = "NP 5 ",
        //            Actions = new List<Namen>()
        //            {
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 5 - Namen 1",
        //                    OpisKolicine = "FO 5 - OK 1",
        //                    Velikost = 5.0M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 50.00M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 5 - Namen 2",
        //                    OpisKolicine = "FO 5 - OK 2",
        //                    Velikost = 5.1M,
        //                    Oznaka ="kom",
        //                    VisinaSpodbude = 50.10M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 5 - Namen 3",
        //                    OpisKolicine = "FO 5 - OK 3",
        //                    Velikost = 5.2M,
        //                    Oznaka ="kW",
        //                    VisinaSpodbude = 50.20M
        //                },
        //                new Namen()
        //                {
        //                    NazivNamena = "FO 5 - Namen 4",
        //                    OpisKolicine = "FO 5 - OK 4",
        //                    Velikost = 5.3M,
        //                    Oznaka ="m2",
        //                    VisinaSpodbude = 50.30M
        //                }
        //            }
        //        }
        //    });
        //    var time2 = DateTime.Now.ToFileTime().ToString();
        //    File.WriteAllBytes($"C:\\test\\Atena.Documents\\SUB-ListOfRecipient_{time2}.docx", listOfRecipient_WordGenerator.Generate());
        //}

        #endregion
    }
}
