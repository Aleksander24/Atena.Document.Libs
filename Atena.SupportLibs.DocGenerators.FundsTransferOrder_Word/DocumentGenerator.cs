using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.Drawing;
using Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word.Models;
using System.Collections.Generic;
using System.Linq;

namespace Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "DemoTest_FundsTransferOrder";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _tenderNumber;
        string _investment;
        string _recipient;
        string _transferOrder;
        string _transferOrderBox;
        string _dateTransfer;
        string _amountTransfer;
        string _contractValue;
        decimal _contractValues;
        string _subtract;
        decimal _subtracts;
        string _responsiblePerson1;
        string _responsiblePerson2;
        string _possiblePayment;
        string _possibleIncentive;
        string _possibleNotify;
        byte[] _faximile;
        List<TableTenderData> _tableTenderDatas;
        #endregion

        #region DocumentGenerator
        public DocumentGenerator(
            string aTenderNumber,
            string aInvestment,
            string aRecipient,
            string aTransferOrder,
            string aTransferOrderBox,
            string aDateTransfer,
            string aAmountTransfer,
            string aContractValue,
            decimal aContractValues,
            string aSubtract,
            decimal aSubtracts,
            string aResponsiblePerson1,
            string aResponsiblePerson2,
            string aPossiblePayment,
            string aPossibleIncentive,
            string aPossibleNotify,
            byte[] aFaximile,
            List<TableTenderData> aTableTenderDatas)
        {
            _tenderNumber = aTenderNumber;
            _investment = aInvestment;
            _recipient = aRecipient;
            _transferOrder = aTransferOrder;
            _transferOrderBox = aTransferOrderBox;
            _dateTransfer = aDateTransfer;
            _amountTransfer = aAmountTransfer;
            _contractValue = aContractValue;
            _contractValues = aContractValues;
            _subtract = aSubtract;
            _subtracts = aSubtracts;
            _responsiblePerson1 = aResponsiblePerson1;
            _responsiblePerson2 = aResponsiblePerson2;
            _possiblePayment = aPossiblePayment;
            _possibleIncentive = aPossibleIncentive;
            _possibleNotify = aPossibleNotify;
            _tableTenderDatas = aTableTenderDatas;
            _faximile = aFaximile;
        }
        #endregion
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

            paragraph = SetImageEkoLogo(section);
            paragraph = SetTenderNumber(section);
            paragraph = SetInvestmentEnded(section);
            paragraph = SetRecipient(section);
            paragraph = SetTransferOrder(section);
            SetTransferOrderBox(document, paragraph);

            SetTenderTable(section);
            paragraph = SetDateTransfer(section);
            SetDateTransferBox(document, paragraph);
            paragraph = SetAmountTransfer(section);
            SetAmountTranserBox(document, paragraph);

            paragraph = SetContractValue(section);
            SetDecimalContractValue(paragraph);
            paragraph = SetSubtract(section);
            SetSubtractDecimal(paragraph);
            paragraph = SetPosiblePaymentIncentive(section);
            paragraph = SetResponsiblePersons(section);
            paragraph = SetPossibleNotify(section);
            paragraph = SetImageSignature(section, _faximile);


            #region Saving document to stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }

        private static IWParagraph SetImageSignature(IWSection section, byte[] aFaximile)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            //FileStream imageStream2 = new FileStream($"C:\\Users\\Aleksanderv\\source\\repos\\Aleksander24\\Atena.Document.Libs\\Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word\\Images\\ekoskladSignature.png", FileMode.Open, FileAccess.Read);
            IWPicture imageSignature = paragraph.AppendPicture(aFaximile);
            imageSignature.TextWrappingStyle = TextWrappingStyle.Square;
            imageSignature.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            imageSignature.VerticalAlignment = ShapeVerticalAlignment.Bottom;
            imageSignature.Width = 280;
            imageSignature.Height = 110;
            return paragraph;
        }

        private IWParagraph SetPossibleNotify(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangePosibleNotify = paragraph.AppendText(_possibleNotify) as WTextRange;
            textRangePosibleNotify.CharacterFormat.FontName = "Calibri";
            textRangePosibleNotify.CharacterFormat.FontSize = 10f;
            return paragraph;
        }

        private IWParagraph SetResponsiblePersons(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangeResponsiblePerson = paragraph.AppendText(_responsiblePerson1) as WTextRange;
            textRangeResponsiblePerson.CharacterFormat.FontName = "Calibri";
            textRangeResponsiblePerson.CharacterFormat.FontSize = 12f;

            //paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            //paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            textRangeResponsiblePerson = paragraph.AppendText(_responsiblePerson2) as WTextRange;
            textRangeResponsiblePerson.CharacterFormat.FontName = "Calibri";
            textRangeResponsiblePerson.CharacterFormat.FontSize = 12f;
            return paragraph;
        }

        private IWParagraph SetPosiblePaymentIncentive(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            WTextRange textRangePosiblePayment = paragraph.AppendText(_possiblePayment) as WTextRange;
            textRangePosiblePayment.CharacterFormat.FontName = "Calibri";
            textRangePosiblePayment.CharacterFormat.FontSize = 9f;
            textRangePosiblePayment.CharacterFormat.Bold = true;

            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            WTextRange textRangePosibleIncentive = paragraph.AppendText(_possibleIncentive) as WTextRange;
            textRangePosibleIncentive.CharacterFormat.FontName = "Calibri";
            textRangePosibleIncentive.CharacterFormat.FontSize = 9f;
            textRangePosibleIncentive.CharacterFormat.Bold = true;
            return paragraph;
        }

        private void SetAmountTranserBox(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox2 = document.AddParagraphStyle("Znesek nakazila") as WParagraphStyle;
            styleBox2.CharacterFormat.FontName = "Calibri";
            styleBox2.CharacterFormat.FontSize = 14f;
            styleBox2.ParagraphFormat.BeforeSpacing = 0;
            styleBox2.ParagraphFormat.AfterSpacing = 0;
            styleBox2.ParagraphFormat.LineSpacing = 10f;
            styleBox2.CharacterFormat.Bold = true;

            IWTextBox textBox2 = paragraph.AppendTextBox(160, 25);
            textBox2.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox2.TextBoxFormat.FillColor = Color.Yellow;
            WParagraph textBoxParagraph = textBox2.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Znesek nakazila");
            textBoxParagraph.AppendText(decimal.Subtract(_contractValues, _subtracts).ToString() + "€");
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
        }

        private void SetSubtractDecimal(IWParagraph paragraph)
        {
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeDifferences = (WTextRange)paragraph.AppendText(_subtracts.ToString() + "€" + Environment.NewLine);
            textRangeDifferences.CharacterFormat.FontName = "Calibri";
            textRangeDifferences.CharacterFormat.FontSize = 12f;
        }

        private IWParagraph SetSubtract(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeDifference = paragraph.AppendText(_subtract) as WTextRange;
            textRangeDifference.CharacterFormat.FontSize = 12f;
            textRangeDifference.CharacterFormat.FontName = "Calibri";
            return paragraph;
        }

        private void SetDecimalContractValue(IWParagraph paragraph)
        {
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeContractValues = (WTextRange)paragraph.AppendText(_contractValues.ToString() + "€");
            textRangeContractValues.CharacterFormat.FontSize = 12f;
            textRangeContractValues.CharacterFormat.FontName = "Calibri";
        }

        private IWParagraph SetContractValue(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            WTextRange textRangeContractValue = paragraph.AppendText(_contractValue) as WTextRange;
            textRangeContractValue.CharacterFormat.FontName = "Calibri";
            textRangeContractValue.CharacterFormat.FontSize = 12f;
            return paragraph;
        }

        private IWParagraph SetAmountTransfer(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeAmountTransfer = paragraph.AppendText(_amountTransfer) as WTextRange;
            textRangeAmountTransfer.CharacterFormat.FontName = "Calibri";
            textRangeAmountTransfer.CharacterFormat.FontSize = 14f;
            return paragraph;
        }

        private static void SetDateTransferBox(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox1 = document.AddParagraphStyle("Datum nakazila") as WParagraphStyle;
            styleBox1.CharacterFormat.FontName = "Calibri";
            styleBox1.CharacterFormat.FontSize = 14f;
            styleBox1.ParagraphFormat.BeforeSpacing = 0;
            styleBox1.ParagraphFormat.AfterSpacing = 0;
            styleBox1.ParagraphFormat.LineSpacing = 10f;
            styleBox1.CharacterFormat.Bold = true;

            IWTextBox textBox1 = paragraph.AppendTextBox(160, 25);
            textBox1.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox1.TextBoxFormat.FillColor = Color.White;
            WParagraph textBoxParagraph = textBox1.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Datum nakazila");
            DateTime aDateTime = DateTime.UtcNow;
            textBoxParagraph.AppendText(aDateTime.ToString("dd.M.yyyy"));
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
        }

        private IWParagraph SetDateTransfer(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeDateTransfer = paragraph.AppendText(_dateTransfer) as WTextRange;
            textRangeDateTransfer.CharacterFormat.FontSize = 14f;
            textRangeDateTransfer.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }

        private void SetTenderTable(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            #region string for Headings in table
            string Tender = "Razpis:";
            string RecipientFunds = "Prejemnik sredstev:";
            string TaxNumber = "Davčna številka: ";
            string AddressRecipient = "Naslov:";
            string ContractNumber = "Številka pogodbe:";
            string TRRForTransfer = "TRR za nakazilo:";
            #endregion
            IWTable tableTender = section.AddTable();
            tableTender.ResetCells(6, 2);
            tableTender.TableFormat.Paddings.All = 2;
            tableTender[0, 0].Width = 250f;
            tableTender[1, 0].Width = 250f;
            tableTender[2, 0].Width = 250f;
            tableTender[3, 0].Width = 250f;
            tableTender[4, 0].Width = 250f;
            tableTender[5, 0].Width = 250f;
            IWTextRange textRangeTableTender = tableTender[0, 0].AddParagraph().AppendText(Tender);
            textRangeTableTender.CharacterFormat.FontSize = 14f;
            textRangeTableTender = tableTender[1, 0].AddParagraph().AppendText(RecipientFunds);
            textRangeTableTender.CharacterFormat.FontSize = 14f;
            textRangeTableTender = tableTender[2, 0].AddParagraph().AppendText(TaxNumber);
            textRangeTableTender.CharacterFormat.FontSize = 14f;
            textRangeTableTender = tableTender[3, 0].AddParagraph().AppendText(AddressRecipient);
            textRangeTableTender.CharacterFormat.FontSize = 14f;
            textRangeTableTender = tableTender[4, 0].AddParagraph().AppendText(ContractNumber);
            textRangeTableTender.CharacterFormat.FontSize = 14f;
            textRangeTableTender = tableTender[5, 0].AddParagraph().AppendText(TRRForTransfer);
            textRangeTableTender.CharacterFormat.FontSize = 14f;

            for (int i = 0; i < _tableTenderDatas.Count; i++)
            {
                tableTender[0, i + 1].Width = 250f;
                textRangeTableTender = tableTender[0, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].Razpis.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;
                tableTender[1, i + 1].Width = 250f;
                textRangeTableTender = tableTender[1, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].PrejemnikSredstev.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;
                tableTender[2, i + 1].Width = 250f;
                textRangeTableTender = tableTender[2, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].DavcnaStevilka.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;
                tableTender[3, i + 1].Width = 250f;
                textRangeTableTender = tableTender[3, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].Naslov.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;
                tableTender[4, i + 1].Width = 250f;
                textRangeTableTender = tableTender[4, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].StevilkaPogodbe.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;
                tableTender[5, i + 1].Width = 250f;
                textRangeTableTender = tableTender[5, i + 1].AddParagraph().AppendText(_tableTenderDatas[i].TRRZaNakazilo.ToString());
                textRangeTableTender.CharacterFormat.FontSize = 14f;

                // alignment.right;
                WTable table = section.Tables[0] as WTable;
                foreach (WTableRow row in table.Rows)
                {
                    foreach (WTableCell cell in row.Cells)
                    {
                        foreach (WParagraph paragraph1 in cell.Paragraphs)
                        {
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].Razpis.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].PrejemnikSredstev.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].DavcnaStevilka.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].Naslov.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].StevilkaPogodbe.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                            if (paragraph1.Text.Contains(_tableTenderDatas[i].TRRZaNakazilo.ToString()))
                                paragraph1.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
                        }
                    }
                }
            }
        }

        private void SetTransferOrderBox(WordDocument document, IWParagraph paragraph)
        {
            WParagraphStyle styleBox = document.AddParagraphStyle("Številka naloga") as WParagraphStyle;
            styleBox.CharacterFormat.FontName = "Calibri";
            styleBox.CharacterFormat.FontSize = 14f;
            styleBox.ParagraphFormat.BeforeSpacing = 0;
            styleBox.ParagraphFormat.AfterSpacing = 0;
            styleBox.ParagraphFormat.LineSpacing = 10f;
            styleBox.CharacterFormat.Bold = true;

            
            IWTextBox textBox = paragraph.AppendTextBox(140, 25);
            textBox.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox.TextBoxFormat.FillColor = Color.White;
            WParagraph textBoxParagraph = textBox.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Številka naloga");
            textBoxParagraph.AppendText(_transferOrderBox);
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
        }

        private IWParagraph SetTransferOrder(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            WTextRange textRangeTransferOrder = paragraph.AppendText(_transferOrder) as WTextRange;
            textRangeTransferOrder.CharacterFormat.FontName = "Calibri";
            textRangeTransferOrder.CharacterFormat.FontSize = 14f;
            textRangeTransferOrder.CharacterFormat.Bold = true;
            return paragraph;
        }

        private IWParagraph SetRecipient(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeRecipient = paragraph.AppendText(_recipient) as WTextRange;
            textRangeRecipient.CharacterFormat.FontName = "Calibri";
            textRangeRecipient.CharacterFormat.FontSize = 14f;
            textRangeRecipient.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }

        private IWParagraph SetInvestmentEnded(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRangeInvestmentEnded = paragraph.AppendText(_investment) as WTextRange;
            textRangeInvestmentEnded.CharacterFormat.FontName = "Calibri";
            textRangeInvestmentEnded.CharacterFormat.FontSize = 14f;
            textRangeInvestmentEnded.CharacterFormat.Bold = true;
            return paragraph;
        }

        private IWParagraph SetTenderNumber(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRangeTenderNumber = paragraph.AppendText(_tenderNumber) as WTextRange;
            textRangeTenderNumber.CharacterFormat.FontName = "Calibri";
            textRangeTenderNumber.CharacterFormat.FontSize = 9f;
            textRangeTenderNumber.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }

        private static IWParagraph SetImageEkoLogo(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            FileStream imageStream = new FileStream($"C:\\Users\\Aleksanderv\\source\\repos\\Aleksander24\\Atena.Document.Libs\\Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word\\Images\\EkoLogo.png", FileMode.Open, FileAccess.Read);
            IWPicture EkoLogo = paragraph.AppendPicture(imageStream);
            EkoLogo.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            EkoLogo.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            EkoLogo.Width = 200;
            EkoLogo.Height = 70;
            return paragraph;
        }
    }
}
