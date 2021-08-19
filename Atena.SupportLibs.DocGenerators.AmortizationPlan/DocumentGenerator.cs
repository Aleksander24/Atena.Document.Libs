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
        List<TableInformationData> _tableInformationDatas;
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


        public DocumentGenerator (
            List<TableAmortizationData> aTableAmortizationDatas,
            List<TableInformationData> aTableInformationDatas,
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
            string aTitleAssumptionsNotesName
            )
        {
            _tableAmortizationDatas = aTableAmortizationDatas;
            _tableInformationDatas = aTableInformationDatas;
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
            EkoLogo.Width = 150;
            EkoLogo.Height = 80;
            EkoLogo.LockAspectRatio = true;

            //aMainTitleAmortizationName
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeMainTitleAmortizationName = paragraph.AppendText(_mainTitleAmortizationName) as WTextRange;
            textRangeMainTitleAmortizationName.CharacterFormat.FontName = "Calibri";
            textRangeMainTitleAmortizationName.CharacterFormat.FontSize = 14f;
            textRangeMainTitleAmortizationName.CharacterFormat.TextColor = Color.Black;
            textRangeMainTitleAmortizationName.CharacterFormat.Bold = true;

            //aLabelContractName
            paragraph = section.AddParagraph(); 
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeLabelContractName = paragraph.AppendText(_labelContractName) as WTextRange;
            textRangeLabelContractName.CharacterFormat.FontName = "Calibri";
            textRangeLabelContractName.CharacterFormat.FontSize = 9f;
            textRangeLabelContractName.CharacterFormat.TextColor = Color.Black;

            //aPartyName
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangePartyName = paragraph.AppendText(_partyName) as WTextRange;
            textRangePartyName.CharacterFormat.FontName = "Calibri";
            textRangePartyName.CharacterFormat.FontSize = 9f;
            textRangePartyName.CharacterFormat.TextColor = Color.Black;

            // tableInformationDataAmortizationPlan
            IWTable tableInformationDataAmortizationPlan = section.AddTable();
            tableInformationDataAmortizationPlan.ResetCells(6, 2);
            tableInformationDataAmortizationPlan.TableFormat.Paddings.All = 2;
            tableInformationDataAmortizationPlan[0, 0].Width = 250f;
            tableInformationDataAmortizationPlan[1, 0].Width = 250f;
            tableInformationDataAmortizationPlan[2, 0].Width = 250f;
            tableInformationDataAmortizationPlan[3, 0].Width = 250f;
            tableInformationDataAmortizationPlan[4, 0].Width = 250f;
            tableInformationDataAmortizationPlan[5, 0].Width = 250f;
           
            IWTextRange textRangeTableInformationDataAmortizationPlan = 
                tableInformationDataAmortizationPlan[0, 0].AddParagraph().AppendText(_loanValueName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan = 
                tableInformationDataAmortizationPlan[1, 0].AddParagraph().AppendText(_ageOfReturnLoanName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[2, 0].AddParagraph().AppendText(_moratoriumName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[3, 0].AddParagraph().AppendText(_interestRateName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[4, 0].AddParagraph().AppendText(_typeOfCalculationName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            textRangeTableInformationDataAmortizationPlan= 
                tableInformationDataAmortizationPlan[5, 0].AddParagraph().AppendText(_firstDateLoanPaidName);
            textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;

            for (int i = 0; i < _tableInformationDatas.Count; i++)
            {
                tableInformationDataAmortizationPlan[0, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[0, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].LoanValue.ToString("C", CultureInfo.CurrentCulture));
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
                tableInformationDataAmortizationPlan[1, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[1, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].AgeOfReturnLoan.ToString());
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
                tableInformationDataAmortizationPlan[2, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[2, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].Moratorium.ToString());
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
                tableInformationDataAmortizationPlan[3, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[3, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].InterestRate.ToString());
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
                tableInformationDataAmortizationPlan[4, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[4, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].TypeOfCalculationLoan.ToString());
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
                tableInformationDataAmortizationPlan[5, i + 1].Width = 250f;
                textRangeTableInformationDataAmortizationPlan = tableInformationDataAmortizationPlan[5, i + 1].AddParagraph().AppendText(_tableInformationDatas[i].FirstDateLoanPaid.ToString("dd/MM/yyyy"));
                textRangeTableInformationDataAmortizationPlan.CharacterFormat.FontSize = 14f;
            }

            // aTitleAssumptionsNotesName
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeTitleAssumptionsNotesName = paragraph.AppendText(_titleAssumptionsNotesName) as WTextRange;
            textRangeTitleAssumptionsNotesName.CharacterFormat.FontName = "Calibri";
            textRangeTitleAssumptionsNotesName.CharacterFormat.FontSize = 14f;
            textRangeTitleAssumptionsNotesName.CharacterFormat.TextColor = Color.Black;
            textRangeTitleAssumptionsNotesName.CharacterFormat.Bold = true;



            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
        }
    }
}
