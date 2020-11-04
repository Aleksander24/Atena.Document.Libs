using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Atena.SupportLibs.DocGenerators.ListOfTransactions_Word.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace Atena.SupportLibs.DocGenerators.ListOfTransactions_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "DemoTest_ListOfTransactions";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _date;
        string _responsiblePerson;
        string _responsiblePerson2;
        List<TableRowsData> _tableRowsDatas;
        string _sumTransactions;
        byte[] _logo;
        string _numberTransferTable;
        string _tenderTable;
        string _recipientTransferTable;
        string _taxNumberTable;
        string _addressTable;
        string _numberContractTable;
        string _amountContractTable;
        string _subtractTable;
        string _trrTable;
        string _amountTransferTable;
        #endregion

        #region DocumentGenerator
        public DocumentGenerator(
            string aDate,
            string aResponsiblePerson,
            string aResponsiblePerson2,
            List<TableRowsData> aTableRowsDatas,
            string aSumTransactions,
            byte[] aLogo,
            string aNumberTransferTable,
            string aTenderTable,
            string aRecipientTransferTable,
            string aTaxNumberTable,
            string aAddressTable,
            string aNumberContractTable,
            string aAmountContractTable,
            string aSubtractTable,
            string aTrrTable,
            string aAmountTransferTable)
        {
            _date = aDate;
            _responsiblePerson = aResponsiblePerson;
            _responsiblePerson2 = aResponsiblePerson2;
            _tableRowsDatas = aTableRowsDatas;
            _sumTransactions = aSumTransactions;
            _logo = aLogo;
            _numberContractTable = aNumberContractTable;
            _tenderTable = aTenderTable;
            _recipientTransferTable = aRecipientTransferTable;
            _taxNumberTable = aTaxNumberTable;
            _addressTable = aAddressTable;
            _numberTransferTable = aNumberTransferTable;
            _amountContractTable = aAmountContractTable;
            _subtractTable = aSubtractTable;
            _trrTable = aTrrTable;
            _amountTransferTable = aAmountTransferTable;
        }
        #endregion
        public byte[] Generate()
        {
            #region Creating document, add section, paragraph, style
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();
            section.PageSetup.Orientation = PageOrientation.Landscape;
            section.PageSetup.Margins.All = 72;

            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 9f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10f;
            style.CharacterFormat.TextColor = Color.Black;

            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion


            paragraph = SetImageEkoLogo(section, _logo);
            paragraph = SetDateTimeText(section);
            paragraph = SetDateTimeBox(document, section);

            
            IWTable table1;
            SetCreateMainTable(section, out paragraph, out table1);
           

            IWTextRange textBold = SetHeadingsMainTable(table1);


            // RowsData = MainTable
            for (int i = 0; i < _tableRowsDatas.Count; i++)
            {
                table1[i + 1, 0].Width = 35f;
                table1[i + 1, 0].AddParagraph().AppendText(_tableRowsDatas[i].StNakazila.ToString());
                table1[i + 1, 1].Width = 80f;
                table1[i + 1, 1].AddParagraph().AppendText(_tableRowsDatas[i].Razpis.ToString());
                table1[i + 1, 2].Width = 100f;
                textBold = table1[i + 1, 2].AddParagraph().AppendText(_tableRowsDatas[i].PrejemnikNakazila.ToString());
                textBold.CharacterFormat.Bold = true;
                table1[i + 1, 3].Width = 65f;
                textBold = table1[i + 1, 3].AddParagraph().AppendText(_tableRowsDatas[i].DavcnaStevilka.ToString());
                textBold.CharacterFormat.Bold = true;
                table1[i + 1, 4].Width = 100f;
                table1[i + 1, 4].AddParagraph().AppendText(_tableRowsDatas[i].Naslov.ToString());
                table1[i + 1, 5].Width = 60f;
                table1[i + 1, 5].AddParagraph().AppendText(_tableRowsDatas[i].StevPogodbe.ToString());
                table1[i + 1, 6].Width = 50f;
                table1[i + 1, 6].AddParagraph().AppendText(_tableRowsDatas[i].ZnesekPogodbe.ToString());
                table1[i + 1, 7].Width = 50f;
                table1[i + 1, 7].AddParagraph().AppendText(_tableRowsDatas[i].Razlika.ToString());
                table1[i + 1, 8].Width = 100f;
                textBold = table1[i + 1, 8].AddParagraph().AppendText(_tableRowsDatas[i].TRR.ToString());
                textBold.CharacterFormat.Bold = true;
                table1[i + 1, 9].Width = 60f;
                textBold = table1[i + 1, 9].AddParagraph().AppendText(_tableRowsDatas[i].ZnesekNakazila.ToString());
                textBold.CharacterFormat.Bold = true;
            }

            SetSumTransactions(section, out paragraph, out textBold);

            paragraph = SetResponsiblePerson1(section);
            paragraph = SetResponsiblePerson2(section);


            #region Saving document to stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion

        }

        private static void SetCreateMainTable(IWSection section, out IWParagraph paragraph, out IWTable table1)
        {
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            table1 = section.AddTable();
            table1.ResetCells(3, 10);
            table1.TableFormat.BackColor = Color.White;
            table1.TableFormat.Paddings.All = 2;
            table1.TableFormat.HorizontalAlignment = RowAlignment.Left;
        }

        private IWTextRange SetHeadingsMainTable(IWTable table1)
        {
            table1[0, 0].Width = 35f;
            IWTextRange textBold = table1[0, 0].AddParagraph().AppendText(_numberTransferTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 1].Width = 80f;
            textBold = table1[0, 1].AddParagraph().AppendText(_tenderTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 2].Width = 100f;
            textBold = table1[0, 2].AddParagraph().AppendText(_recipientTransferTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 3].Width = 65f;
            textBold = table1[0, 3].AddParagraph().AppendText(_taxNumberTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 4].Width = 100f;
            textBold = table1[0, 4].AddParagraph().AppendText(_addressTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 5].Width = 60f;
            textBold = table1[0, 5].AddParagraph().AppendText(_numberContractTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 6].Width = 50f;
            textBold = table1[0, 6].AddParagraph().AppendText(_amountContractTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 7].Width = 50f;
            textBold = table1[0, 7].AddParagraph().AppendText(_subtractTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 8].Width = 100f;
            textBold = table1[0, 8].AddParagraph().AppendText(_trrTable);
            textBold.CharacterFormat.Bold = true;
            table1[0, 9].Width = 60f;
            textBold = table1[0, 9].AddParagraph().AppendText(_amountTransferTable);
            textBold.CharacterFormat.Bold = true;
            return textBold;
        }

        private void SetSumTransactions(IWSection section, out IWParagraph paragraph, out IWTextRange textBold)
        {
            paragraph = section.AddParagraph();
            IWTable tableSumNakazila = section.AddTable();
            tableSumNakazila.ResetCells(1, 2);
            tableSumNakazila.TableFormat.HorizontalAlignment = RowAlignment.Right;
            tableSumNakazila.TableFormat.Paddings.All = 2;
            tableSumNakazila[0, 0].Width = 70f;
            textBold = tableSumNakazila[0, 0].AddParagraph().AppendText(_sumTransactions);
            textBold.CharacterFormat.FontSize = 12f;

            decimal sumNakazila = _tableRowsDatas.Sum(p => p.ZnesekNakazila);
            tableSumNakazila[0, 1].Width = 60f;
            textBold = tableSumNakazila[0, 1].AddParagraph().AppendText(sumNakazila.ToString() + "€");
            textBold.CharacterFormat.Bold = true;
            textBold.CharacterFormat.FontSize = 12f;
        }

        private static IWParagraph SetDateTimeBox(WordDocument document, IWSection section)
        {
            IWParagraph paragraph;
            WParagraphStyle styleBox = document.AddParagraphStyle("Datum") as WParagraphStyle;
            styleBox.CharacterFormat.FontName = "Calibri";
            styleBox.CharacterFormat.FontSize = 12f;
            styleBox.ParagraphFormat.BeforeSpacing = 0;
            styleBox.ParagraphFormat.AfterSpacing = 0;
            styleBox.ParagraphFormat.LineSpacing = 10f;
            styleBox.CharacterFormat.Bold = true;

            paragraph = section.AddParagraph();
            IWTextBox textBox = paragraph.AppendTextBox(80, 20);
            textBox.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            textBox.TextBoxFormat.FillColor = Color.White;
            WParagraph textBoxParagraph = textBox.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Datum");
            DateTime aDateTime = DateTime.UtcNow;
            textBoxParagraph.AppendText(aDateTime.ToString("dd.M.yyyy"));
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            return paragraph;
        }

        private IWParagraph SetResponsiblePerson2(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRange3 = paragraph.AppendText(_responsiblePerson2) as WTextRange;
            textRange3.CharacterFormat.FontName = "Calibri";
            textRange3.CharacterFormat.FontSize = 12f;
            textRange3.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }

        private IWParagraph SetResponsiblePerson1(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange2 = paragraph.AppendText(_responsiblePerson) as WTextRange;
            textRange2.CharacterFormat.FontName = "Calibri";
            textRange2.CharacterFormat.FontSize = 12f;
            textRange2.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }

        private IWParagraph SetDateTimeText(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRange1 = paragraph.AppendText(_date) as WTextRange;
            textRange1.CharacterFormat.FontName = "Calibri";
            textRange1.CharacterFormat.FontSize = 10f;
            return paragraph;
        }

        private static IWParagraph SetImageEkoLogo(IWSection section, byte[] logo)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            //FileStream imageStream = new FileStream($"D:\\DeloOdDoma\\Atena.Document.Libs\\Atena.SupportLibs.DocGenerators.ListOfTransactions_Word\\Image\\EkoLogo.png", FileMode.Open, FileAccess.Read);
            IWPicture EkoLogo = paragraph.AppendPicture(logo);
            EkoLogo.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
            EkoLogo.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            EkoLogo.Width = 200;
            EkoLogo.Height = 70;
            return paragraph;
        }
    }
}
