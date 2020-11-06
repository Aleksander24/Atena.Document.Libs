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
        string _borrowerTextBorrowerBox;
        string _investNameTextBorrowerBox;
        string _amountCreditBorrowerBoxTable;
        string _contractBorrowerBoxTable; 
        string _maturityRepayBorrowerBoxTable; 
        string _dateSignatureBorrowerBoxTable; 
        string _moratoriumBorrowerBoxTable;
        string _personTextCreatedReportBox; 
        string _nameSurnameTextCreatedReportBox; 
        string _functionTextCreatedReportBox; 
        string _phoneFaxTextCreatedReportBox;
        string _investProcent;
        string _levelConditionInvestText;
        string _frontBehindInvestLevelConditionInvestBoxTable; 
        string _endInvestLevelConditionInvestBoxTable; 
        string _year1WorkLevelConditionInvestBoxTable; 
        string _year2WorkLevelConditionInvestBoxTable; 
        string _year3WorkLevelConditionInvestBoxTable;
        string _techDataInvestText;
        string _paramText; 
        string _unitText; 
        string _forecastText; 
        string _realizeText; 
        string _footnoteText;
        string _hoursText; 
        string _usingEnergy;
        List<RowDatasBasicTechInvest> _rowDatasBasicTechInvests;
        string _regularOperationTextBox;
        #endregion

        public DocumentGenerator(
            string aHeadDocumentText,
            string aConcernText,
            string aBorrowerTextBorrowerBox,
            string aInvestNameTextBorrowerBox,
            string aAmountCreditBorrowerBoxTable,
            string aContractBorrowerBoxTable,
            string aMaturityRepayBorrowerBoxTable,
            string aDateSignatureBorrowerBoxTable,
            string aMoratoriumBorrowerBoxTable,
            string aPersonTextCreatedReportBox,
            string aNameSurnameTextCreatedReportBox,
            string aFunctionTextCreatedReportBox,
            string aPhoneFaxTextCreatedReportBox,
            string aInvestProcent,
            string aLevelConditionInvestText,
            string aFrontBehindInvestLevelConditionInvestBoxTable,
            string aEndInvestLevelConditionInvestBoxTable,
            string aYear1WorkLevelConditionInvestBoxTable,
            string aYear2WorkLevelConditionInvestBoxTable,
            string aYear3WorkLevelConditionInvestBoxTable, 
            string aTechDataInvestText, 
            string aParamText,
            string aUnitText,
            string aForecastText,
            string aRealizeText,
            string aFootnoteText,
            string aHoursText, 
            string aUsingEnergy,
            List<RowDatasBasicTechInvest> aRowDatasBasicTechInvests,
            string aRegularOperationTextBox
            )
        {
            _headDocumentText = aHeadDocumentText;
            _concernText = aConcernText;
            _borrowerTextBorrowerBox = aBorrowerTextBorrowerBox;
            _investNameTextBorrowerBox = aInvestNameTextBorrowerBox;
            _amountCreditBorrowerBoxTable = aAmountCreditBorrowerBoxTable;
            _contractBorrowerBoxTable = aContractBorrowerBoxTable;
            _maturityRepayBorrowerBoxTable = aMaturityRepayBorrowerBoxTable;
            _dateSignatureBorrowerBoxTable = aDateSignatureBorrowerBoxTable;
            _moratoriumBorrowerBoxTable = aMoratoriumBorrowerBoxTable;
            _personTextCreatedReportBox = aPersonTextCreatedReportBox;
            _nameSurnameTextCreatedReportBox = aNameSurnameTextCreatedReportBox;
            _functionTextCreatedReportBox = aFunctionTextCreatedReportBox;
            _phoneFaxTextCreatedReportBox = aPhoneFaxTextCreatedReportBox;
            _investProcent = aInvestProcent;
            _levelConditionInvestText = aLevelConditionInvestText;
            _frontBehindInvestLevelConditionInvestBoxTable = aFrontBehindInvestLevelConditionInvestBoxTable;
            _endInvestLevelConditionInvestBoxTable = aEndInvestLevelConditionInvestBoxTable;
            _year1WorkLevelConditionInvestBoxTable = aYear1WorkLevelConditionInvestBoxTable;
            _year2WorkLevelConditionInvestBoxTable = aYear2WorkLevelConditionInvestBoxTable;
            _year3WorkLevelConditionInvestBoxTable = aYear3WorkLevelConditionInvestBoxTable;
            _techDataInvestText = aTechDataInvestText;
            _paramText = aParamText;
            _unitText = aUnitText;
            _forecastText = aForecastText;
            _realizeText = aRealizeText;
            _footnoteText = aFootnoteText;
            _hoursText = aHoursText;
            _usingEnergy = aUsingEnergy;
            _rowDatasBasicTechInvests = aRowDatasBasicTechInvests;
            _regularOperationTextBox = aRegularOperationTextBox;
        }

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
            paragraph = section.AddParagraph();
            IWTextBox textBox1 = SetLegalPersonBox2(document, paragraph);
            WTextRange textRangeInsideBorrower = SetLegalPersonBoxTable(textBox1);
            IWTextBox textBox2;
            AddPersonCreatedReportBox3(section, out paragraph, out textBox2);
            textRangeInsideBorrower = SetPersonCreatedReportBoxTable(textBox2);
            paragraph = SetInvestProcentTable(section);
            IWTextBox textBoxLevelConditionInvest3 = AddLevelConditionInvestBox4(document, paragraph);
            paragraph = SetLevelConditionInvestBoxTable(section, textBoxLevelConditionInvest3);

            paragraph = SetTechDataInvestText(section);
            IWTextBox textBoxTechDataInvest4;
            AddBasicTechDataInvestBox5(section, out paragraph, out textBoxTechDataInvest4);
            SetBasicTexhDataInvestBoxTable(textBoxTechDataInvest4);

            // string regularOperationTextBox
            paragraph = section.AddParagraph();
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

            #region Saving document to stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }

        private void SetBasicTexhDataInvestBoxTable(IWTextBox textBoxTechDataInvest4)
        {
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
            textRangeTechDataInvest = (WTextRange)tableTechDataInvestHeads[0, 4].AddParagraph().AppendText(_footnoteText);
            textRangeTechDataInvest.CharacterFormat.FontSize = 10f;
            textRangeTechDataInvest.CharacterFormat.Bold = true;

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
            textBoxTechDataInvest4 = paragraph.AppendTextBox(500, 70);
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

        private IWParagraph SetLevelConditionInvestBoxTable(IWSection section, IWTextBox textBoxLevelConditionInvest3)
        {
            IWParagraph paragraph = section.AddParagraph();
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
            return paragraph;
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

            IWTextBox textBoxLevelConditionInvest3 = paragraph.AppendTextBox(350, 83);
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
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsidePersonCreatedReport[0, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[0, 1].AddParagraph().AppendText(_nameSurnameTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsidePersonCreatedReport[1, 0].Width = 216.6f;
            tableInsidePersonCreatedReport[1, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[1, 1].AddParagraph().AppendText(_functionTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsidePersonCreatedReport[2, 0].Width = 216.6f;
            tableInsidePersonCreatedReport[2, 1].Width = 283.4f;
            textRangeInsideBorrower = (WTextRange)tableInsidePersonCreatedReport[2, 1].AddParagraph().AppendText(_phoneFaxTextCreatedReportBox);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            return textRangeInsideBorrower;
        }

        private static void AddPersonCreatedReportBox3(IWSection section, out IWParagraph paragraph, out IWTextBox textBox2)
        {
            paragraph = section.AddParagraph();
            textBox2 = paragraph.AppendTextBox(500, 64);
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
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsideBorrower1[0, 1].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[0, 1].AddParagraph().AppendText(_contractBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsideBorrower1[0, 2].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[0, 2].AddParagraph().AppendText(_maturityRepayBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsideBorrower1[1, 0].Width = 166.6f;
            tableInsideBorrower1[1, 1].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[1, 1].AddParagraph().AppendText(_dateSignatureBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            tableInsideBorrower1[1, 2].Width = 166.6f;
            textRangeInsideBorrower = (WTextRange)tableInsideBorrower1[1, 2].AddParagraph().AppendText(_moratoriumBorrowerBoxTable);
            textRangeInsideBorrower.CharacterFormat.FontSize = 12f;
            return textRangeInsideBorrower;
        }

        private IWTextBox SetLegalPersonBox2(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox2 = document.AddParagraphStyle("Box pravna oseba") as WParagraphStyle;
            styleBox2.CharacterFormat.FontName = "Calibri";
            styleBox2.CharacterFormat.FontSize = 12f;
            styleBox2.ParagraphFormat.BeforeSpacing = 0;
            styleBox2.ParagraphFormat.AfterSpacing = 0;
            styleBox2.ParagraphFormat.LineSpacing = 10f;

            IWTextBox textBox1 = paragraph.AppendTextBox(500, 83);
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
            IWTextBox textBoxConcern = paragraph.AppendTextBox(200, 20);
            textBoxConcern.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBoxConcern.TextBoxFormat.FillColor = Color.White;
        }

        private IWParagraph SetConcernText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeConcern = paragraph.AppendText(_concernText) as WTextRange;
            textRangeConcern.CharacterFormat.FontName = "Calibri";
            textRangeConcern.CharacterFormat.FontSize = 12f;
            return paragraph;
        }

        private IWParagraph SetHeadDocument(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangeHeadDoc = paragraph.AppendText(_headDocumentText) as WTextRange;
            textRangeHeadDoc.CharacterFormat.FontName = "Calibri";
            textRangeHeadDoc.CharacterFormat.FontSize = 16f;
            textRangeHeadDoc.CharacterFormat.Bold = true;
            return paragraph;
        }
    }
}
