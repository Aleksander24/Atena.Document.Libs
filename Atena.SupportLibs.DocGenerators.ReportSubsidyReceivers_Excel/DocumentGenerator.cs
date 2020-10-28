using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using Syncfusion.XlsIO;
using Atena.SupportLibs.DocGenerators.ReportSubsidyReceivers_Excel.Models;
using System.Globalization;

namespace Atena.SupportLibs.DocGenerators.ReportSubsidyReceivers_Excel
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";
        public string Label => "DemoTest_ReportSubsidyReceivers";
        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Excel;

        #region PRIVATS
        string _receiver;
        string _addressReceiver;
        string _mailID;
        string _taxNumber;
        string _parameterDesc;
        string _amountHelp;
        string _dateDesicion;
        List<RowsData> _rowDatas;
        #endregion
        public DocumentGenerator(
            string aReceiver, 
            string aAddressReceiver, 
            string aMailID, 
            string aTaxNumber,
            string aParameterDesc, 
            string aAmountHelp, 
            string aDateDesicion,
            List<RowsData> aRowDatas)
        {
            _receiver = aReceiver;
            _addressReceiver = aAddressReceiver;
            _mailID = aMailID;
            _taxNumber = aTaxNumber;
            _parameterDesc = aParameterDesc;
            _amountHelp = aAmountHelp;
            _dateDesicion = aDateDesicion;
            _rowDatas = aRowDatas;

        }
        public byte[] Generate()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                #region Creating Excel version and workbook
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2010; // set up confirming excel

                // Create workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                #endregion

                // rows data
                for (int i = 0; i < _rowDatas.Count; i++)
                {
                    worksheet[i + 1, 1].Text = _rowDatas[i].Prejemnik.ToString();
                    worksheet[i + 1, 2].Text = _rowDatas[i].NaslovPrejemnika.ToString();
                    worksheet[i + 1, 3].Text = _rowDatas[i].PostaID.ToString();
                    worksheet[i + 1, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                    worksheet[i + 1, 4].Number = _rowDatas[i].DavcnaStevilka;
                    worksheet[i + 1, 5].Text = _rowDatas[i].OpisParametra.ToString();
                    worksheet[i + 1, 6].Number = (double)_rowDatas[i].VisinaPomoci;
                    worksheet[i + 1, 6].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                    worksheet[i + 1, 6].NumberFormat = "#.##€";
                    worksheet[i + 1, 7].Text = _rowDatas[i].DatumOdlocbe.ToString();
                    worksheet[i + 1, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                    worksheet[i + 1, 7].NumberFormat = "dd.mm.yyyy";
                }

                #region Entering values to the cells = Heading
                // Inserting another row ((for) start with 0)
                worksheet.InsertRow(1, 1, ExcelInsertOptions.FormatAsBefore);
                worksheet.Range["A1"].Text = _receiver;
                worksheet.Range["B1"].Text = _addressReceiver;
                worksheet.Range["C1"].Text = _mailID;
                worksheet.Range["D1"].Text = _taxNumber;
                worksheet.Range["E1"].Text = _parameterDesc;
                worksheet.Range["F1"].Text = _amountHelp;
                worksheet.Range["G1"].Text = _dateDesicion;
                worksheet.Range["A1:G1"].CellStyle.Font.Bold = true;
                #endregion

                #region Saving excel document
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                return stream.ToArray();
                #endregion
            }
        }
    }
}
