﻿using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Collections.Generic;
using Syncfusion.Drawing;
using Atena.SupportLibs.DocGenerators.SUB_ListOfRecipient_Word.Models;
using System.Linq;
using System.Globalization;

namespace Atena.SupportLibs.DocGenerators.SUB_ListOfRecipient_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "DemoTest_SUB-ListOfRecipient";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _textFinancialIncentive;
        string _textPayouts;
        List<ReceiverData> _rowDatas;
        string _serialNumberText;
        string _recipientIncentiveText;
        string _addressRecipientText;
        string _purposeText;
        string _descriptionQuantityText;
        string _heightText;
        string _unitText;
        string _amountIncentiveText;
        #endregion
        
        #region DocumentGenerator
        public DocumentGenerator(
            string aTextFinancialIncentive, 
            string aTextPayouts, 
            List<ReceiverData> aRowDatas, 
            string aSerialNumberText,
            string aRecipientIncentiveText,
            string aAddressRecipientText,
            string aPurposeText,
            string aDescriptionQuantityText,
            string aHeightText,
            string aUnitText,
            string aAmountIncentiveText)
        {
            _textFinancialIncentive = aTextFinancialIncentive;
            _textPayouts = aTextPayouts;
            _rowDatas = aRowDatas;
            _serialNumberText = aSerialNumberText;
            _recipientIncentiveText = aRecipientIncentiveText;
            _addressRecipientText = aAddressRecipientText;
            _purposeText = aPurposeText;
            _descriptionQuantityText = aDescriptionQuantityText;
            _heightText = aHeightText;
            _unitText = aUnitText;
            _amountIncentiveText = aAmountIncentiveText;
        }
        #endregion

        public byte[] Generate()
        {
            #region Creating word, add section, add style
            //Creating a new document
            WordDocument document = new WordDocument();

            //Adding a new section to the document
            IWSection section = document.AddSection();

            //Set Margin of the section
            section.PageSetup.Margins.All = 40;

            //Set page size of the section
            section.PageSetup.PageSize = new SizeF(575, 792);

            //Create Paragraph styles for normal font
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 9f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10f;
            style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;

            //Create Paragraph styles for heading font
            style = document.AddParagraphStyle("Heading 1") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 14f;
            style.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.Keep = true;
            style.ParagraphFormat.KeepFollow = true;
            style.ParagraphFormat.OutlineLevel = OutlineLevel.Level1;
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion

            SetTextFinancialIncentive(_textFinancialIncentive, section);

            SetTextPayouts(_textPayouts, section);

            IWTable table = CreatingTable(section);
            
            SetHeadingsForTable(table);

            // rows in table SUBListOfRecipient
            for (int i = 0; i < _rowDatas.Count; i++)
            {
                table[i + 1, 0].Width = 26f;
                table[i + 1, 0].AddParagraph().AppendText(_rowDatas[i].ZapStevilka.ToString());
                table[i + 1, 1].Width = 70f;
                table[i + 1, 1].AddParagraph().AppendText(_rowDatas[i].PrejemnikSpodbude.ToString());
                table[i + 1, 2].Width = 70f;
                table[i + 1, 2].AddParagraph().AppendText(_rowDatas[i].NaslovPrejemnika.ToString());

                for (int j = 0; j < _rowDatas[i].Actions.Count; j++)
                {
                    table[i + 1, 3].Width = 160f;
                    table[i + 1, 3].AddParagraph().AppendText(_rowDatas[i].Actions[j].NazivNamena.ToString() + Environment.NewLine);
                    table[i + 1, 4].Width = 80f;
                    table[i + 1, 4].AddParagraph().AppendText(_rowDatas[i].Actions[j].OpisKolicine.ToString() + Environment.NewLine);
                    table[i + 1, 5].Width = 40f;
                    table[i + 1, 5].AddParagraph().AppendText(_rowDatas[i].Actions[j].Velikost.ToString() + Environment.NewLine); 
                    table[i + 1, 6].Width = 25f;
                    table[i + 1, 6].AddParagraph().AppendText(_rowDatas[i].Actions[j].Oznaka.ToString() + Environment.NewLine); 
                    table[i + 1, 7].Width = 45f;
                    table[i + 1, 7].AddParagraph().AppendText(_rowDatas[i].Actions[j].VisinaSpodbude.ToString() + Environment.NewLine); 
                }
            }

            #region Saving word document
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;

            return stream.ToArray();
            #endregion
        }

        private static IWParagraph SetTextPayouts(string aTextPayouts, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRange2 = paragraph.AppendText(aTextPayouts) as WTextRange;
            textRange2.CharacterFormat.FontSize = 11f;
            textRange2.CharacterFormat.FontName = "Calibri";
            textRange2.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }
        private static IWParagraph SetTextFinancialIncentive(string aTextFinancialIncentive, IWSection section)
            {
                IWParagraph paragraph = section.AddParagraph();
                paragraph.ApplyStyle("Normal");
                paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
                WTextRange textRange1 = paragraph.AppendText(aTextFinancialIncentive) as WTextRange;
                textRange1.CharacterFormat.FontSize = 11f;
                textRange1.CharacterFormat.FontName = "Calibri";
                textRange1.CharacterFormat.TextColor = Color.Black;
                textRange1.CharacterFormat.Bold = true;
                return paragraph;
            }
        private static IWTable CreatingTable(IWSection section)
        {
            IWTable table = section.AddTable();
            table.ResetCells(6, 8);
            table.TableFormat.BackColor = Color.White;
            table.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table.TableFormat.Positioning.HorizPosition = 1;
            table.TableFormat.Positioning.VertPosition = 90;
            table.TableFormat.Paddings.All = 2;
            return table;
        }
        private void SetHeadingsForTable(IWTable table)
        {
            table[0, 0].Width = 26f;
            IWTextRange tabletextRange = table[0, 0].AddParagraph().AppendText(_serialNumberText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 1].Width = 70f;
            tabletextRange = table[0, 1].AddParagraph().AppendText(_recipientIncentiveText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 2].Width = 70f;
            tabletextRange = table[0, 2].AddParagraph().AppendText(_addressRecipientText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 3].Width = 160f;
            tabletextRange = table[0, 3].AddParagraph().AppendText(_purposeText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 4].Width = 80f;
            tabletextRange = table[0, 4].AddParagraph().AppendText(_descriptionQuantityText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 5].Width = 40f;
            tabletextRange = table[0, 5].AddParagraph().AppendText(_heightText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 6].Width = 25f;
            tabletextRange = table[0, 6].AddParagraph().AppendText(_unitText);
            tabletextRange.CharacterFormat.Bold = true;
            table[0, 7].Width = 45f;
            tabletextRange = table[0, 7].AddParagraph().AppendText(_amountIncentiveText);
            tabletextRange.CharacterFormat.Bold = true;
        }
    }
}
