using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Atena.SupportLibs.DocGenerators.AmortizationPlan.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Atena.SupportLibs.DocGenerators.AmortizationPlan
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";
        public string Label => "DemoTest_AmortizationPlan";
        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        List<TableAmortizationData> _tableAmortizationDatas;
        //List<TableInformationData> _tableInformationDatas;
        string _mainTitleAmortizationName; // glavni naslov dokumenta
        byte[] _logo; // slika na desnem zgornjem vrhu
        string _labelContractName; // Oznaka pogodbe:
        string _partyName; // Partija:
        string _loanValueName; // Vrednost kredita:
        string _ageOfReturnLoanName; // Doba vračanja:
        string _moratoriumName; // Moratorij:
        string _interestRateName; // Obrestna mera:
        string _typeOfCalculationName; // Način obračuna:
        string _firstDateLoanPaidName; // datum prvega obroka:
        string _titleAssumptionsNotesName; // PREDPOSTAVKE IN OPOMBE:

        string _labelContractValue;
        string _partyValue;         
        double _loanValue; 
        int _ageOfReturnLoanNumber; 
        int _moratoriumNumber;         
        double _interestRateValue; 
        string _typeOfCalculation; 
        DateTime _firstDateLoanPaid = new DateTime();

        public DocumentGenerator (
            List<TableAmortizationData> aTableAmortizationDatas,
            string aMainTitleAmortizationName,
            byte[] aLogo,
            string aLabelContractName,
            string aPartyName,
            string aLoanValueName,
            string aAgeOfReturnLoanName,
            string aMoratoriumName,
            string aInterestRateName,
            string aTypeOfCalculationName,
            string aFirstDateLoanPaidName,
            string aTitleAssumptionsNotesName,
            string aLabelContractValue,
            string aPartyValue,
            double aLoanValue,
            int aAgeOfReturnLoanNumber,
            int aMoratoriumNumber,
            double aInterestRateValue,
            string aTypeOfCalculation,
            DateTime aFirstDateLoanPaid = new DateTime()
            )
        {
            _tableAmortizationDatas = aTableAmortizationDatas;
            //_tableInformationDatas = aTableInformationDatas;
            _mainTitleAmortizationName = aMainTitleAmortizationName;
            _logo = aLogo;
            _labelContractName = aLabelContractName;
            _partyName = aPartyName;
            _loanValueName = aLoanValueName;
            _ageOfReturnLoanName = aAgeOfReturnLoanName;
            _moratoriumName = aMoratoriumName;
            _interestRateName = aInterestRateName;
            _typeOfCalculationName = aTypeOfCalculationName;
            _firstDateLoanPaidName = aFirstDateLoanPaidName;
            _titleAssumptionsNotesName = aTitleAssumptionsNotesName;
            _labelContractValue = aLabelContractValue;
            _partyValue = aPartyValue;
            _loanValue = aLoanValue;
            _ageOfReturnLoanNumber = aAgeOfReturnLoanNumber;
            _moratoriumNumber = aMoratoriumNumber;
            _interestRateValue = aInterestRateValue;
            _typeOfCalculation = aTypeOfCalculation;
            _firstDateLoanPaid = aFirstDateLoanPaid;

        }
        public byte[] Generate()
        {
            #region Creating document, adding section, edit style, paragraph
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();
            section.PageSetup.Margins.All = 40;
            section.PageSetup.PageSize = new SizeF(575, 792);

            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 10f;
            style.CharacterFormat.TextColor = Color.Black;

            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion

            //aLogo
            paragraph = section.HeadersFooters.Header.AddParagraph();
            section.PageSetup.HeaderDistance = 10f;
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            paragraph.ApplyStyle("Normal");
            WPicture EkoLogo = paragraph.AppendPicture(_logo) as WPicture;
            EkoLogo.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            EkoLogo.Width = 120;
            EkoLogo.Height = 80;
            EkoLogo.LockAspectRatio = true;

            //aMainTitleAmortizationName
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeMainTitleAmortizationName = paragraph.AppendText(_mainTitleAmortizationName) as WTextRange;
            textRangeMainTitleAmortizationName.CharacterFormat.FontName = "Calibri";
            textRangeMainTitleAmortizationName.CharacterFormat.FontSize = 16f;
            textRangeMainTitleAmortizationName.CharacterFormat.TextColor = Color.Black;
            textRangeMainTitleAmortizationName.CharacterFormat.Bold = true;

            #region tableInformationDataAmortizationPlan
            paragraph = section.AddParagraph();
            IWTable tableInformationDataAmortizationPlan = section.AddTable();
            tableInformationDataAmortizationPlan.ResetCells(8, 2);
            tableInformationDataAmortizationPlan.TableFormat.Paddings.All = 2;
            tableInformationDataAmortizationPlan[0, 0].Width = 250f;
            tableInformationDataAmortizationPlan[1, 0].Width = 250f;
            tableInformationDataAmortizationPlan[2, 0].Width = 250f;
            tableInformationDataAmortizationPlan[3, 0].Width = 250f;
            tableInformationDataAmortizationPlan[4, 0].Width = 250f;
            tableInformationDataAmortizationPlan[5, 0].Width = 250f;
            tableInformationDataAmortizationPlan[6, 0].Width = 250f;
            tableInformationDataAmortizationPlan[7, 0].Width = 250f;
            string mesec = " mes.";
            string procent = " %";
            IWTextRange textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[0, 0].AddParagraph().AppendText(_labelContractName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[1, 0].AddParagraph().AppendText(_partyName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan = 
                tableInformationDataAmortizationPlan[2, 0].AddParagraph().AppendText(_loanValueName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan = 
                tableInformationDataAmortizationPlan[3, 0].AddParagraph().AppendText(_ageOfReturnLoanName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[4, 0].AddParagraph().AppendText(_moratoriumName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[5, 0].AddParagraph().AppendText(_interestRateName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[6, 0].AddParagraph().AppendText(_typeOfCalculationName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[7, 0].AddParagraph().AppendText(_firstDateLoanPaidName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;

            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[0, 1].AddParagraph().AppendText(_labelContractValue);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[1, 1].AddParagraph().AppendText(_partyValue);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[2, 1].AddParagraph().AppendText(_loanValue.ToString("C", CultureInfo.CurrentCulture));
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[3, 1].AddParagraph().AppendText(_ageOfReturnLoanNumber.ToString() + mesec);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[4, 1].AddParagraph().AppendText(_moratoriumNumber.ToString() + mesec);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[5, 1].AddParagraph().AppendText(_interestRateValue.ToString() + procent);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[6, 1].AddParagraph().AppendText(_typeOfCalculation);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan =
                tableInformationDataAmortizationPlan[7, 1].AddParagraph().AppendText(_firstDateLoanPaid.ToString("d/M/yyyy"));
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            #endregion

            // aTitleAssumptionsNotesName
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeTitleAssumptionsNotesName = paragraph.AppendText(_titleAssumptionsNotesName) as WTextRange;
            textRangeTitleAssumptionsNotesName.CharacterFormat.FontName = "Calibri";
            textRangeTitleAssumptionsNotesName.CharacterFormat.FontSize = 14f;
            textRangeTitleAssumptionsNotesName.CharacterFormat.TextColor = Color.Black;
            textRangeTitleAssumptionsNotesName.CharacterFormat.Bold = true;


            paragraph = section.AddParagraph();
            IWTable tableAmortizationData = section.AddTable();
            tableAmortizationData.ResetCells(1, 8);
            tableAmortizationData.TableFormat.Paddings.All = 2;

            IWTextRange textRangetableAmortizationData =
                tableAmortizationData[0, 0].AddParagraph().AppendText("Meseci");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 0].Width = 50f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 1].AddParagraph().AppendText("Anuiteta");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 1].Width = 50f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 2].AddParagraph().AppendText("Plačane obresti");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 2].Width = 70f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 3].AddParagraph().AppendText("Plačana glavnica");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 3].Width = 80f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 4].AddParagraph().AppendText("Bilanca");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 4].Width = 80f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 5].AddParagraph().AppendText("Datum nakazila");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 5].Width = 60f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 6].AddParagraph().AppendText("EURIBOR+OM");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 6].Width = 75f;
            textRangetableAmortizationData =
                tableAmortizationData[0, 7].AddParagraph().AppendText("Leto");
            textRangetableAmortizationData.CharacterFormat.FontSize = 12f;
            tableAmortizationData[0, 7].Width = 30f;

            int i = 1;
            
            foreach (var amortizationData in _tableAmortizationDatas)
            {
                int j = 1;
                //var date = _tableAmortizationDatas[i].StartLoanDate;
                WTableRow tableRow = tableAmortizationData.AddRow(true);

                tableRow.Cells[0].AddParagraph().AppendText(i.ToString());
                tableRow.Cells[1].AddParagraph().AppendText(amortizationData.Annuity.ToString());
                tableRow.Cells[2].AddParagraph().AppendText(amortizationData.InterestPaid.ToString("C"));
                tableRow.Cells[3].AddParagraph().AppendText(amortizationData.MonthlyPay.ToString("C"));
                tableRow.Cells[4].AddParagraph().AppendText(amortizationData.Balance.ToString("C"));
                tableRow.Cells[5].AddParagraph().AppendText(amortizationData.StartLoanDate.AddMonths(j).ToString("d.M.yyyy"));
                tableRow.Cells[6].AddParagraph().AppendText(amortizationData.Euribor_OM.ToString() + procent);
                tableRow.Cells[7].AddParagraph().AppendText(amortizationData.StartLoanDate.Year.ToString());
                i++;
            }

            //for (int i = 0; i < _tableInformationDatas.Count; i++)
            //{
            //    tableInformationDataAmortizationPlan[0, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[0, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].LoanValue.ToString("C", CultureInfo.CurrentCulture));
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //    tableInformationDataAmortizationPlan[1, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[1, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].AgeOfReturnLoan.ToString());
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //    tableInformationDataAmortizationPlan[2, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[2, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].Moratorium.ToString());
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //    tableInformationDataAmortizationPlan[3, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[3, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].InterestRate.ToString());
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //    tableInformationDataAmortizationPlan[4, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[4, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].TypeOfCalculationLoan.ToString());
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //    tableInformationDataAmortizationPlan[5, i + 1].Width = 250f;
            //    textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[5, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].FirstDateLoanPaid.ToString("dd/MM/yyyy"));
            //    textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            //}

            



            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
        }
    }
}
