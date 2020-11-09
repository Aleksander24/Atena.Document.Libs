using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.IO;
using Atena.SupportLibs.DocGenerators.ReportInvestmentEffects_Word.Models;
using System.Collections.Generic;

namespace Atena.SupportLibs.DocGenerators.ReportInvestmentEffects_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "TestDemo_ReportInvestmentEffects_Word";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _headDocumentText;
        string _concernText;
        #region BorrowerTextBox
        string _borrowerTextBorrowerBox;
        string _investNameTextBorrowerBox;
        string _amountCreditBorrowerBoxTable;
        string _contractBorrowerBoxTable; 
        string _maturityRepayBorrowerBoxTable; 
        string _dateSignatureBorrowerBoxTable; 
        string _moratoriumBorrowerBoxTable;
        #endregion
        #region ResponsiblePersonCreatedReportTextBox
        string _personTextCreatedReportBox; 
        string _nameSurnameTextCreatedReportBox; 
        string _functionTextCreatedReportBox; 
        string _phoneFaxTextCreatedReportBox;
        #endregion
        string _investProcent;
        #region LevelConditionSituationInvestTextBox
        string _levelConditionInvestText;
        string _frontBehindInvestLevelConditionInvestBoxTable; 
        string _endInvestLevelConditionInvestBoxTable; 
        string _year1WorkLevelConditionInvestBoxTable; 
        string _year2WorkLevelConditionInvestBoxTable; 
        string _year3WorkLevelConditionInvestBoxTable;
        #endregion
        string _techDataInvestText;
        #region TechDataInvestTextBoxTable
        string _paramText; 
        string _unitText; 
        string _forecastText; 
        string _realizeText; 
        string _footnoteTextTechDataInvest;
        string _hoursText; 
        string _usingEnergy;
        List<RowDatasBasicTechInvest> _rowDatasBasicTechInvests;
        #endregion
        string _regularOperationTextBox;
        string _performEffects;
        #region TableHeadPerformEffects
        string _paramPerformEffectsHead; 
        string _unitPerformEffects; 
        string _situatInvestPerformEffectsHead; 
        string _forecastPerformEffectsHead; 
        string _year1PerformEffectsHead; 
        string _year2PerformEffectsHead; 
        string _year3PerformEffectsHead;
        string _footnotePerformEffectsHead;
        List<RowDatasPerformEffects> _rowDatasPerformEffects;
        #endregion
        string _emergencyInstruction;
        string _footNoteText;
        string _dateCreatedReportText;
        string _signatureReportText;
        string _generalInstructionsHeadText;
        string _generalInstructionsData;
        #endregion

        #region DocumentGenerator
        public DocumentGenerator(
            string aHeadDocumentText,
            string aConcernText,
        #region BorrowerTextBox
            string aBorrowerTextBorrowerBox,
            string aInvestNameTextBorrowerBox,
            string aAmountCreditBorrowerBoxTable,
            string aContractBorrowerBoxTable,
            string aMaturityRepayBorrowerBoxTable,
            string aDateSignatureBorrowerBoxTable,
            string aMoratoriumBorrowerBoxTable,
        #endregion
        #region ResponsiblePersonCreatedReportTextBox
            string aPersonTextCreatedReportBox,
            string aNameSurnameTextCreatedReportBox,
            string aFunctionTextCreatedReportBox,
            string aPhoneFaxTextCreatedReportBox,
        #endregion
            string aInvestProcent,
        #region LevelConditionSituationInvestTextBox
            string aLevelConditionInvestText,
            string aFrontBehindInvestLevelConditionInvestBoxTable,
            string aEndInvestLevelConditionInvestBoxTable,
            string aYear1WorkLevelConditionInvestBoxTable,
            string aYear2WorkLevelConditionInvestBoxTable,
            string aYear3WorkLevelConditionInvestBoxTable,
        #endregion
            string aTechDataInvestText,
        #region TechDataInvestTextBoxTable
            string aParamText,
            string aUnitText,
            string aForecastText,
            string aRealizeText,
            string aFootnoteTextTechDataInvest,
            string aHoursText,
            string aUsingEnergy,
            List<RowDatasBasicTechInvest> aRowDatasBasicTechInvests,
        #endregion
            string aRegularOperationTextBox,
            string aPerformEffects,
        #region TableHeadPerformEffects
            string aParamPerformEffectsHead,
            string aUnitPerformEffects,
            string aSituatInvestPerformEffectsHead,
            string aforecastPerformEffectsHead,
            string aYear1PerformEffectsHead,
            string aYear2PerformEffectsHead,
            string aYear3PerformEffectsHead,
            string aFootnotePerformEffectsHead,
            List<RowDatasPerformEffects> aRowDatasPerformEffects,
        #endregion
            string aEmergencyInstruction,
            string aFootNoteText,
            string aDateCreatedReportText,
            string aSignatureReportText,
            string aGeneralInstructionsHeadText,
            string aGeneralInstructionsData
            )
        {
            _headDocumentText = aHeadDocumentText;
            _concernText = aConcernText;
            #region BorrowerTextBox
            _borrowerTextBorrowerBox = aBorrowerTextBorrowerBox;
            _investNameTextBorrowerBox = aInvestNameTextBorrowerBox;
            _amountCreditBorrowerBoxTable = aAmountCreditBorrowerBoxTable;
            _contractBorrowerBoxTable = aContractBorrowerBoxTable;
            _maturityRepayBorrowerBoxTable = aMaturityRepayBorrowerBoxTable;
            _dateSignatureBorrowerBoxTable = aDateSignatureBorrowerBoxTable;
            _moratoriumBorrowerBoxTable = aMoratoriumBorrowerBoxTable;
            #endregion
            #region ResponsiblePersonCreatedReportTextBox
            _personTextCreatedReportBox = aPersonTextCreatedReportBox;
            _nameSurnameTextCreatedReportBox = aNameSurnameTextCreatedReportBox;
            _functionTextCreatedReportBox = aFunctionTextCreatedReportBox;
            _phoneFaxTextCreatedReportBox = aPhoneFaxTextCreatedReportBox;
            #endregion
            _investProcent = aInvestProcent;
            #region LevelConditionSituationInvestTextBox
            _levelConditionInvestText = aLevelConditionInvestText;
            _frontBehindInvestLevelConditionInvestBoxTable = aFrontBehindInvestLevelConditionInvestBoxTable;
            _endInvestLevelConditionInvestBoxTable = aEndInvestLevelConditionInvestBoxTable;
            _year1WorkLevelConditionInvestBoxTable = aYear1WorkLevelConditionInvestBoxTable;
            _year2WorkLevelConditionInvestBoxTable = aYear2WorkLevelConditionInvestBoxTable;
            _year3WorkLevelConditionInvestBoxTable = aYear3WorkLevelConditionInvestBoxTable;
            #endregion
            #region TechDataInvestTextBoxTable
            _techDataInvestText = aTechDataInvestText;
            _paramText = aParamText;
            _unitText = aUnitText;
            _forecastText = aForecastText;
            _realizeText = aRealizeText;
            _footnoteTextTechDataInvest = aFootnoteTextTechDataInvest;
            _hoursText = aHoursText;
            _usingEnergy = aUsingEnergy;
            _rowDatasBasicTechInvests = aRowDatasBasicTechInvests;
            #endregion
            _regularOperationTextBox = aRegularOperationTextBox;
            _performEffects = aPerformEffects;
            _paramPerformEffectsHead = aParamPerformEffectsHead;
            #region TableHeadPerformEffects
            _unitPerformEffects = aUnitPerformEffects;
            _situatInvestPerformEffectsHead = aSituatInvestPerformEffectsHead;
            _forecastPerformEffectsHead = aforecastPerformEffectsHead;
            _year1PerformEffectsHead = aYear1PerformEffectsHead;
            _year2PerformEffectsHead = aYear2PerformEffectsHead;
            _year3PerformEffectsHead = aYear3PerformEffectsHead;
            _footnotePerformEffectsHead = aFootnotePerformEffectsHead;
            #endregion
            _rowDatasPerformEffects = aRowDatasPerformEffects;
            _emergencyInstruction = aEmergencyInstruction;
            _footNoteText = aFootNoteText;
            _dateCreatedReportText = aDateCreatedReportText;
            _signatureReportText = aSignatureReportText;
            _generalInstructionsHeadText = aGeneralInstructionsHeadText;
            _generalInstructionsData = aGeneralInstructionsData;
        }
        #endregion
        public byte[] Generate()
        {
            #region Creating document, add paragraph, section, style
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

            paragraph = SetHeadDocument(section);
            paragraph = SetConcernText(section);
            AddConcernBox1(paragraph);

            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            IWTextBox textBox1 = SetLegalPersonBox2(document, paragraph);
            WTextRange textRangeInsideBorrower = SetLegalPersonBoxTable(textBox1);
            IWTextBox textBox2;
            AddPersonCreatedReportBox3(section, out paragraph, out textBox2);
            textRangeInsideBorrower = SetPersonCreatedReportBoxTable(textBox2);
            paragraph = SetInvestProcentTable(section);
            IWTextBox textBoxLevelConditionInvest3 = AddLevelConditionInvestBox4(document, paragraph);
            SetLevelConditionInvestBoxTable(textBoxLevelConditionInvest3);

            paragraph = SetTechDataInvestText(section);
            IWTextBox textBoxTechDataInvest4;
            AddBasicTechDataInvestBox5(section, out paragraph, out textBoxTechDataInvest4);
            SetBasicTechDataInvestBoxTable(textBoxTechDataInvest4);
            paragraph = SetRegularTextBox6(document, section);
            paragraph = SetPerformEffectsText(section);
            IWTextBox textBoxPerformEffects5;
            AddPerformEffectsTextBox7(section, out paragraph, out textBoxPerformEffects5);
            SetPerformEffectsTextBoxTable(textBoxPerformEffects5);
            paragraph = SetEmergencyInstructionText(section);
            paragraph = SetFootNoteText(section);
            paragraph = AddFootNoteTextBox8(section);
            paragraph = SetSignatureReportAndCreatedReportText(section);
            WTextBox textBoxGeneralInstructions;
            AddGeneralInstructionsTextbBox9(section, out paragraph, out textBoxGeneralInstructions);
            SetGeneralInstructionsText(paragraph, textBoxGeneralInstructions);

            #region Saving document to stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }

        private void SetGeneralInstructionsText(IWParagraph paragraph, WTextBox textBoxGeneralInstructions)
        {
            WTextRange textRangeGeneralInstructionsHead = textBoxGeneralInstructions.TextBoxBody.AddParagraph().AppendText(_generalInstructionsHeadText) as WTextRange;
            textRangeGeneralInstructionsHead.CharacterFormat.FontSize = 12f;
            textRangeGeneralInstructionsHead.CharacterFormat.Bold = true;

            //paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeGeneralInstructionsData = textBoxGeneralInstructions.TextBoxBody.AddParagraph().AppendText(_generalInstructionsData) as WTextRange;
            textRangeGeneralInstructionsData.CharacterFormat.FontSize = 10f;
            textRangeGeneralInstructionsData.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
        }

        private static void AddGeneralInstructionsTextbBox9(IWSection section, out IWParagraph paragraph, out WTextBox textBoxGeneralInstructions)
        {
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            textBoxGeneralInstructions = paragraph.AppendTextBox(420, 75) as WTextBox;
            textBoxGeneralInstructions.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            textBoxGeneralInstructions.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
            textBoxGeneralInstructions.TextBoxFormat.FillColor = Color.White;
        }

        private IWParagraph SetSignatureReportAndCreatedReportText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeDateCreatedReport = paragraph.AppendText(_dateCreatedReportText) as WTextRange;
            textRangeDateCreatedReport.CharacterFormat.FontSize = 10f;

            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeSignatureReport = paragraph.AppendText(_signatureReportText) as WTextRange;
            textRangeSignatureReport.CharacterFormat.FontSize = 10f;
            return paragraph;
        }

        private static IWParagraph AddFootNoteTextBox8(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            WTextBox textBoxFootNote = paragraph.AppendTextBox(500, 15) as WTextBox;
            textBoxFootNote.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxFootNote.TextBoxFormat.FillColor = Color.White;
            textBoxFootNote.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
            return paragraph;
        }

        private IWParagraph SetFootNoteText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeFootNoteText = paragraph.AppendText(_footNoteText) as WTextRange;
            textRangeFootNoteText.CharacterFormat.FontSize = 12f;
            textRangeFootNoteText.CharacterFormat.Bold = true;
            textRangeFootNoteText.CharacterFormat.Italic = true;
            return paragraph;
        }

        private IWParagraph SetEmergencyInstructionText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangeEmergencyInstruction = paragraph.AppendText(_emergencyInstruction) as WTextRange;
            textRangeEmergencyInstruction.CharacterFormat.FontSize = 12f;
            textRangeEmergencyInstruction.CharacterFormat.TextColor = Color.Red;
            textRangeEmergencyInstruction.CharacterFormat.TextBackgroundColor = Color.Yellow;
            textRangeEmergencyInstruction.CharacterFormat.Italic = true;
            textRangeEmergencyInstruction.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
            return paragraph;
        }

        private void SetPerformEffectsTextBoxTable(IWTextBox textBoxPerformEffects5)
        {
            #region HeadsTable
            WTable tablePerformEffectsHeads = textBoxPerformEffects5.TextBoxBody.AddTable() as WTable;
            tablePerformEffectsHeads.ResetCells(1, 8);
            tablePerformEffectsHeads.TableFormat.Paddings.All = 2;
            tablePerformEffectsHeads.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tablePerformEffectsHeads[0, 0].Width = 145f;
            WTextRange textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 0].AddParagraph().AppendText(_paramPerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 1].Width = 30f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 1].AddParagraph().AppendText(_unitPerformEffects);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 2].Width = 70f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 2].AddParagraph().AppendText(_situatInvestPerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 3].Width = 50f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 3].AddParagraph().AppendText(_forecastPerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 4].Width = 35f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 4].AddParagraph().AppendText(_year1PerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 5].Width = 35f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 5].AddParagraph().AppendText(_year2PerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 6].Width = 35f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 6].AddParagraph().AppendText(_year3PerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            tablePerformEffectsHeads[0, 7].Width = 100f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsHeads[0, 7].AddParagraph().AppendText(_footnotePerformEffectsHead);
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            #endregion
            #region Rows
            WTable tablePerformEffectsRows = textBoxPerformEffects5.TextBoxBody.AddTable() as WTable;
            tablePerformEffectsRows.ResetCells(2, 9);
            tablePerformEffectsRows.TableFormat.Paddings.All = 2;
            tablePerformEffectsRows.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tablePerformEffectsRows[0, 0].Width = 25f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsRows[0, 0].AddParagraph().AppendText(_hoursText);
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            textRangePerformEffects.CharacterFormat.Italic = true;
            textRangePerformEffects.CharacterFormat.TextColor = Color.Red;
            textRangePerformEffects.CharacterFormat.Bold = true;
            tablePerformEffectsRows[0, 1].Width = 120f;
            textRangePerformEffects = (WTextRange)tablePerformEffectsRows[0, 1].AddParagraph().AppendText(_usingEnergy);
            textRangePerformEffects.CharacterFormat.FontSize = 10f;
            textRangePerformEffects.CharacterFormat.Italic = true;
            textRangePerformEffects.CharacterFormat.TextColor = Color.Red;
            textRangePerformEffects.CharacterFormat.Bold = true;
            tablePerformEffectsRows[0, 2].Width = 30f;
            tablePerformEffectsRows[0, 3].Width = 70f;
            tablePerformEffectsRows[0, 4].Width = 50f;
            tablePerformEffectsRows[0, 5].Width = 35f;
            tablePerformEffectsRows[0, 6].Width = 35f;
            tablePerformEffectsRows[0, 7].Width = 35f;
            tablePerformEffectsRows[0, 8].Width = 100f;
            #endregion
            // class rowdata
            for (int i = 0; i < _rowDatasPerformEffects.Count; i++)
            {
                tablePerformEffectsRows[i + 1, 0].Width = 25f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 0].AddParagraph().AppendText(_rowDatasPerformEffects[i].Ure.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 1].Width = 120f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 1].AddParagraph().AppendText(_rowDatasPerformEffects[i].RabaEnergije.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 2].Width = 30f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 2].AddParagraph().AppendText(_rowDatasPerformEffects[i].Enota.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 3].Width = 70f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 3].AddParagraph().AppendText(_rowDatasPerformEffects[i].StanjePredInvesticijo.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 4].Width = 50f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 4].AddParagraph().AppendText(_rowDatasPerformEffects[i].Prognoza.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 5].Width = 35f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 5].AddParagraph().AppendText(_rowDatasPerformEffects[i].Leto1.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 6].Width = 35f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 6].AddParagraph().AppendText(_rowDatasPerformEffects[i].Leto2.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 7].Width = 35f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 7].AddParagraph().AppendText(_rowDatasPerformEffects[i].Leto3.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
                tablePerformEffectsRows[i + 1, 8].Width = 100f;
                textRangePerformEffects = (WTextRange)tablePerformEffectsRows[i + 1, 8].AddParagraph().AppendText(_rowDatasPerformEffects[i].Opomba.ToString());
                textRangePerformEffects.CharacterFormat.FontSize = 9f;
            }
        }

        private static void AddPerformEffectsTextBox7(IWSection section, out IWParagraph paragraph, out IWTextBox textBoxPerformEffects5)
        {
            paragraph = section.AddParagraph();
            textBoxPerformEffects5 = paragraph.AppendTextBox(500, 77);
            textBoxPerformEffects5.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxPerformEffects5.TextBoxFormat.FillColor = Color.White;
            textBoxPerformEffects5.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
        }

        private IWParagraph SetPerformEffectsText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangePerformEffects = paragraph.AppendText(_performEffects) as WTextRange;
            textRangePerformEffects.CharacterFormat.FontSize = 12f;
            textRangePerformEffects.CharacterFormat.Bold = true;
            textRangePerformEffects.CharacterFormat.Italic = true;
            return paragraph;
        }

        private IWParagraph SetRegularTextBox6(WordDocument document, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            WParagraphStyle styleBox5 = document.AddParagraphStyle("Box rednega obratovanja") as WParagraphStyle;
            styleBox5.CharacterFormat.FontName = "Calibri";
            styleBox5.CharacterFormat.FontSize = 9f;
            styleBox5.ParagraphFormat.BeforeSpacing = 0;
            styleBox5.ParagraphFormat.AfterSpacing = 0;
            styleBox5.ParagraphFormat.LineSpacing = 10f;
            WTextBox textBoxRegularOperation = (WTextBox)paragraph.AppendTextBox(150, 17);
            textBoxRegularOperation.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            textBoxRegularOperation.TextBoxFormat.FillColor = Color.White;
            textBoxRegularOperation.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
            WParagraph textBoxParagraph5 = textBoxRegularOperation.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph5.ApplyStyle("Box rednega obratovanja");
            textBoxParagraph5.AppendText(_regularOperationTextBox);
            textBoxParagraph5.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            return paragraph;
        }

        private void SetBasicTechDataInvestBoxTable(IWTextBox textBoxTechDataInvest4)
        {
            #region Heads
            WTable tableTechDataInvestHeads = textBoxTechDataInvest4.TextBoxBody.AddTable() as WTable;
            tableTechDataInvestHeads.ResetCells(1, 5);
            tableTechDataInvestHeads[0, 0].Width = 170f;
            tableTechDataInvestHeads.TableFormat.Paddings.All = 2;
            tableTechDataInvestHeads.TableFormat.HorizontalAlignment = RowAlignment.Center;
            WTextRange textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 0].AddParagraph().AppendText(_paramText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestHeads[0, 1].Width = 50f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 1].AddParagraph().AppendText(_unitText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestHeads[0, 2].Width = 80f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 2].AddParagraph().AppendText(_forecastText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestHeads[0, 3].Width = 90f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 3].AddParagraph().AppendText(_realizeText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestHeads[0, 4].Width = 110f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 4].AddParagraph().AppendText(_footnoteTextTechDataInvest);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            #endregion
            #region DataRows
            WTable tableTechDataInvestRows = textBoxTechDataInvest4.TextBoxBody.AddTable() as WTable;
            tableTechDataInvestRows.ResetCells(2, 6);
            tableTechDataInvestRows.TableFormat.Paddings.All = 2;
            tableTechDataInvestRows.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableTechDataInvestRows[0, 0].Width = 30f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[0, 0].AddParagraph().AppendText(_hoursText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Italic = true;
            textRangeTechDataInvest.CharacterFormat.TextColor = Color.Red;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestRows[0, 1].Width = 140f;
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[0, 1].AddParagraph().AppendText(_usingEnergy);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Italic = true;
            textRangeTechDataInvest.CharacterFormat.TextColor = Color.Red;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            tableTechDataInvestRows[0, 2].Width = 50f;
            tableTechDataInvestRows[0, 3].Width = 80f;
            tableTechDataInvestRows[0, 4].Width = 90f;
            tableTechDataInvestRows[0, 5].Width = 110f;
            #endregion
            for (int i = 0; i < _rowDatasBasicTechInvests.Count; i++)
            {
                tableTechDataInvestRows[i + 1, 0].Width = 30f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 0].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].Ure.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
                tableTechDataInvestRows[i + 1, 1].Width = 140f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 1].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].RabaEnergije.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
                tableTechDataInvestRows[i + 1, 2].Width = 50f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 2].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].Enota.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
                tableTechDataInvestRows[i + 1, 3].Width = 80f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 3].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].Prognoza.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
                tableTechDataInvestRows[i + 1, 4].Width = 90f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 4].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].Realizirano.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
                tableTechDataInvestRows[i + 1, 5].Width = 110f;
                textRangeTechDataInvest = (WTextRange)tableTechDataInvestRows[i + 1, 5].AddParagraph().AppendText(_rowDatasBasicTechInvests[i].Opomba.ToString());
                textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            }
        }

        private static void AddBasicTechDataInvestBox5(IWSection section, out IWParagraph paragraph, out IWTextBox textBoxTechDataInvest4)
        {
            paragraph = section.AddParagraph();
            textBoxTechDataInvest4 = paragraph.AppendTextBox(500, 61);
            textBoxTechDataInvest4.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxTechDataInvest4.TextBoxFormat.FillColor = Color.White;
            textBoxTechDataInvest4.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
        }

        private IWParagraph SetTechDataInvestText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeTechDataInvest = paragraph.AppendText(_techDataInvestText) as WTextRange;
            textRangeTechDataInvest.CharacterFormat.FontSize = 12f;
            textRangeTechDataInvest.CharacterFormat.Italic = true;
            textRangeTechDataInvest.CharacterFormat.Bold = true;
            return paragraph;
        }
        private void SetLevelConditionInvestBoxTable(IWTextBox textBoxLevelConditionInvest3)
        {
            WTable tableLevelConditionInvest = textBoxLevelConditionInvest3.TextBoxBody.AddTable() as WTable;
            tableLevelConditionInvest.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableLevelConditionInvest.TableFormat.Paddings.All = 3;
            tableLevelConditionInvest.ResetCells(3, 2);

            tableLevelConditionInvest[0, 0].Width = 175f;
            WTextRange textRangeLevelConditionInvest = (WTextRange)tableLevelConditionInvest[0, 0].AddParagraph().AppendText(_frontBehindInvestLevelConditionInvestBoxTable + "[   ]");
            textRangeLevelConditionInvest.CharacterFormat.FontSize = 9f;
            tableLevelConditionInvest[0, 1].Width = 175f;
            textRangeLevelConditionInvest = (WTextRange)tableLevelConditionInvest[0, 1].AddParagraph().AppendText(_year1WorkLevelConditionInvestBoxTable + "[   ]");
            textRangeLevelConditionInvest.CharacterFormat.FontSize = 9f;
            tableLevelConditionInvest[1, 0].Width = 175f;
            textRangeLevelConditionInvest = (WTextRange)tableLevelConditionInvest[1, 0].AddParagraph().AppendText(_endInvestLevelConditionInvestBoxTable + "[   ]");
            textRangeLevelConditionInvest.CharacterFormat.FontSize = 9f;
            tableLevelConditionInvest[1, 1].Width = 175f;
            textRangeLevelConditionInvest = (WTextRange)tableLevelConditionInvest[1, 1].AddParagraph().AppendText(_year2WorkLevelConditionInvestBoxTable + "[   ]");
            textRangeLevelConditionInvest.CharacterFormat.FontSize = 9f;
            tableLevelConditionInvest[2, 0].Width = 175f;
            tableLevelConditionInvest[2, 1].Width = 175f;
            textRangeLevelConditionInvest = (WTextRange)tableLevelConditionInvest[2, 1].AddParagraph().AppendText(_year3WorkLevelConditionInvestBoxTable + "[   ]");
            textRangeLevelConditionInvest.CharacterFormat.FontSize = 9f;
        }
        
        private IWTextBox AddLevelConditionInvestBox4(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox3 = document.AddParagraphStyle("Head box Stanje oz.") as WParagraphStyle;
            styleBox3.CharacterFormat.FontName = "Calibri";
            styleBox3.CharacterFormat.FontSize = 12f;
            styleBox3.CharacterFormat.Bold = true;
            styleBox3.CharacterFormat.Italic = true;
            styleBox3.ParagraphFormat.BeforeSpacing = 0;
            styleBox3.ParagraphFormat.AfterSpacing = 0;
            styleBox3.ParagraphFormat.LineSpacing = 10f;

            IWTextBox textBoxLevelConditionInvest3 = paragraph.AppendTextBox(350, 66);
            textBoxLevelConditionInvest3.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxLevelConditionInvest3.TextBoxFormat.FillColor = Color.White;
            WParagraph textBoxParagraph = textBoxLevelConditionInvest3.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Head box Stanje oz.");
            textBoxParagraph.AppendText(_levelConditionInvestText);
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            return textBoxLevelConditionInvest3;
        }

        private IWParagraph SetInvestProcentTable(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            IWTable tableInvestProcent = section.AddTable();
            tableInvestProcent.ResetCells(1, 2);
            tableInvestProcent[0, 0].Width = 85f;
            WTextRange textRangeInvestProcent = (WTextRange)tableInvestProcent[0, 0].AddParagraph().AppendText(_investProcent);
            textRangeInvestProcent.CharacterFormat.FontSize = 12f;
            textRangeInvestProcent.CharacterFormat.Bold = true;
            textRangeInvestProcent.CharacterFormat.Italic = true;
            tableInvestProcent[0, 1].Width = 40f;
            textRangeInvestProcent = (WTextRange)tableInvestProcent[0, 1].AddParagraph().AppendText(Environment.NewLine + "%");
            textRangeInvestProcent.CharacterFormat.FontSize = 14f;
            // % right alignment
            WTable table = section.Tables[0] as WTable;
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    foreach (WParagraph paragraph1 in cell.Paragraphs)
                    {
                        if (paragraph1.Text.Contains("%"))
                            paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;

                    }
                }
            }

            return paragraph;
        }

        private WTextRange SetPersonCreatedReportBoxTable(IWTextBox textBox2)
        {
            WTextRange textRangeInsideBorrower;
            WTable tableInsidePersonCreatedReport = textBox2.TextBoxBody.AddTable() as WTable;
            tableInsidePersonCreatedReport.ResetCells(3, 2);
            tableInsidePersonCreatedReport.TableFormat.BackColor = Color.White;
            tableInsidePersonCreatedReport.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableInsidePersonCreatedReport.TableFormat.Paddings.All = 2;

            tableInsidePersonCreatedReport[0, 0].Width = 216.6f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[0, 0].AddParagraph().AppendText(_personTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsidePersonCreatedReport[0, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[0, 1].AddParagraph().AppendText(_nameSurnameTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsidePersonCreatedReport[1, 0].Width = 216.6f;
            tableInsidePersonCreatedReport[1, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[1, 1].AddParagraph().AppendText(_functionTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsidePersonCreatedReport[2, 0].Width = 216.6f;
            tableInsidePersonCreatedReport[2, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[2, 1].AddParagraph().AppendText(_phoneFaxTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            return textRangeInsideBorrower;
        }

        private static void AddPersonCreatedReportBox3(IWSection section, out IWParagraph paragraph, out IWTextBox textBox2)
        {
            paragraph = section.AddParagraph();
            textBox2 = paragraph.AppendTextBox(500, 56);
            textBox2.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox2.TextBoxFormat.FillColor = Color.White;
            textBox2.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
        }

        private WTextRange SetLegalPersonBoxTable(IWTextBox textBox1)
        {
            WTable tableInsideBorrower1 = textBox1.TextBoxBody.AddTable() as WTable;
            tableInsideBorrower1.ResetCells(2, 3);
            tableInsideBorrower1.TableFormat.BackColor = Color.White;
            tableInsideBorrower1.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableInsideBorrower1.TableFormat.Paddings.All = 2;

            tableInsideBorrower1[0, 0].Width = 166.6f;
            WTextRange textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[0, 0].AddParagraph().AppendText(_amountCreditBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsideBorrower1[0, 1].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[0, 1].AddParagraph().AppendText(_contractBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsideBorrower1[0, 2].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[0, 2].AddParagraph().AppendText(_maturityRepayBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsideBorrower1[1, 0].Width = 166.6f;
            tableInsideBorrower1[1, 1].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[1, 1].AddParagraph().AppendText(_dateSignatureBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            tableInsideBorrower1[1, 2].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[1, 2].AddParagraph().AppendText(_moratoriumBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 10f;
            return textRangeInsideBorrower;
        }

        private IWTextBox SetLegalPersonBox2(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox2 = document.AddParagraphStyle("Box pravna oseba") as WParagraphStyle;
            styleBox2.CharacterFormat.FontName = "Calibri";
            styleBox2.CharacterFormat.FontSize = 10f;
            styleBox2.ParagraphFormat.BeforeSpacing = 0;
            styleBox2.ParagraphFormat.AfterSpacing = 0;
            styleBox2.ParagraphFormat.LineSpacing = 10f;

            IWTextBox textBox1 = paragraph.AppendTextBox(500, 76);
            textBox1.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox1.TextBoxFormat.FillColor = Color.White;
            textBox1.TextBoxFormat.TextWrappingStyle = TextWrappingStyle.TopAndBottom;
            WParagraph textBoxParagraph = textBox1.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Box pravna oseba");
            textBoxParagraph.AppendText(_borrowerTextBorrowerBox + _investNameTextBorrowerBox + Environment.NewLine);
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            return textBox1;
        }

        private static void AddConcernBox1(IWParagraph paragraph)
        {
            IWTextBox textBoxConcern = paragraph.AppendTextBox(200, 15);
            textBoxConcern.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxConcern.TextBoxFormat.FillColor = Color.White;
        }

        private IWParagraph SetConcernText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeConcern = paragraph.AppendText(_concernText) as WTextRange;
            textRangeConcern.CharacterFormat.FontName = "Calibri";
            textRangeConcern.CharacterFormat.FontSize = 11f;
            return paragraph;
        }

        private IWParagraph SetHeadDocument(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangeHeadDoc = paragraph.AppendText(_headDocumentText) as WTextRange;
            textRangeHeadDoc.CharacterFormat.FontName = "Calibri";
            textRangeHeadDoc.CharacterFormat.FontSize = 14f;
            textRangeHeadDoc.CharacterFormat.Bold = true;
            return paragraph;
        }
    }
}
