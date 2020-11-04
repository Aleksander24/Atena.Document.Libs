using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ListOfRemittances_FinishedUnfinished_Word.Models;
using ListOfRemittances_FinishedUnfinished_Word.Models.UnfinishedData;
using ListOfRemittances_FinishedUnfinished_Word.Models.FinishedData;

namespace ListOfRemittances_FinishedUnfinished_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "DemoTest_ListOfRemmittances_FinishedUnfinished";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _head;
        string _dateRemittances;
        List<UnTenderCode1> _unTenderCode1s;
        List<UnTenderCode2> _unTenderCode2s;
        List<UnTenderCode3> _unTenderCode3s;
        List<FiTenderCode1> _fiTenderCode1s;
        List<FiTenderCode2> _fiTenderCode2s;
        List<FiTenderCode3> _fiTenderCode3s;
        string _investStatusText;
        string _finishedText;
        string _unfinishedText;
        string _tenderUnit;
        string _unfinishedTenderNumber1;
        string _unfinishedTenderNumber2;
        string _unfinishedTenderNumber3;
        string _finishedTenderNumber1;
        string _finishedTenderNumber2;
        string _finishedTenderNumber3;
        #endregion

        public DocumentGenerator(
            string aHead,
            string aDateRemittances,
            List<UnTenderCode1> aUnTenderCode1s,
            List<UnTenderCode2> aUnTenderCode2s,
            List<UnTenderCode3> aUnTenderCode3s,
            List<FiTenderCode1> aFiTenderCode1s,
            List<FiTenderCode2> aFiTenderCode2s,
            List<FiTenderCode3> aFiTenderCode3s,
            string aInvestStatusText,
            string aFinishedText,
            string aUnfinishedText,
            string aTenderUnit,
            string aUnfinishedTenderNumber1,
            string aUnfinishedTenderNumber2,
            string aUnfinishedTenderNumber3,
            string aFinishedTenderNumber1,
            string aFinishedTenderNumber2,
            string aFinishedTenderNumber3) 
        {
            _head = aHead;
            _dateRemittances = aDateRemittances;
            _unTenderCode1s = aUnTenderCode1s;
            _unTenderCode2s = aUnTenderCode2s;
            _unTenderCode3s = aUnTenderCode3s;
            _fiTenderCode1s = aFiTenderCode1s;
            _fiTenderCode2s = aFiTenderCode2s;
            _fiTenderCode3s = aFiTenderCode3s;
            _investStatusText = aInvestStatusText;
            _finishedText = aFinishedText;
            _unfinishedText = aUnfinishedText;
            _tenderUnit = aTenderUnit;
            _unfinishedTenderNumber1 = aUnfinishedTenderNumber1;
            _unfinishedTenderNumber2 = aUnfinishedTenderNumber2;
            _unfinishedTenderNumber3 = aUnfinishedTenderNumber3;
            _finishedTenderNumber1 = aFinishedTenderNumber1;
            _finishedTenderNumber2 = aFinishedTenderNumber2;
            _finishedTenderNumber3 = aFinishedTenderNumber3;
        }

        public byte[] Generate()
        {
            #region Creating document, edit style, section and add paragraph
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();
            section.PageSetup.Margins.All = 40;
            section.PageSetup.PageSize = new SizeF(575, 792);

            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 10f;
            style.CharacterFormat.TextColor = Color.Black;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10f;

            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion

            paragraph = SetHeadDocument(section);
            paragraph = SetDateRemittances(section);
            paragraph = SetDateTimeBox(document, section);
            paragraph = SetTableHeadUnfinished(section);

            #region OznakaRazpisa1Unfinished
            IWTable tableUnfinished1 = section.AddTable();
            tableUnfinished1.ResetCells(4, 5);
            tableUnfinished1.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableUnfinished1.TableFormat.BackColor = Color.White;
            tableUnfinished1.TableFormat.Paddings.All = 2;

            tableUnfinished1[0, 0].Width = 28f;
            tableUnfinished1[0, 1].Width = 30f;
            tableUnfinished1[0, 2].Width = 100f;
            WTextRange textRangeUnfinished1 = (WTextRange)tableUnfinished1[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeUnfinished1.CharacterFormat.Bold = true;
            textRangeUnfinished1.CharacterFormat.FontSize = 12f;
            tableUnfinished1[0, 3].Width = 100f;
            textRangeUnfinished1 = (WTextRange)tableUnfinished1[0, 3].AddParagraph().AppendText(_unfinishedTenderNumber1);
            textRangeUnfinished1.CharacterFormat.Bold = true;
            textRangeUnfinished1.CharacterFormat.FontSize = 12f;
            tableUnfinished1[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _unTenderCode1s.Count; i++)
            {
                tableUnfinished1[i + 1, 0].Width = 28f;
                tableUnfinished1[i + 1, 0].AddParagraph().AppendText(_unTenderCode1s[i].ZapStevilka1.ToString());
                tableUnfinished1[i + 1, 1].Width = 30f;
                tableUnfinished1[i + 1, 1].AddParagraph().AppendText(_unTenderCode1s[i].ZapStevilka2.ToString());
                tableUnfinished1[i + 1, 2].Width = 100f;
                tableUnfinished1[i + 1, 2].AddParagraph().AppendText(_unTenderCode1s[i].Oznaka1.ToString());
                tableUnfinished1[i + 1, 3].Width = 100f;
                tableUnfinished1[i + 1, 3].AddParagraph().AppendText(_unTenderCode1s[i].Oznaka2.ToString());
                tableUnfinished1[i + 1, 4].Width = 240f;
                tableUnfinished1[i + 1, 4].AddParagraph().AppendText(_unTenderCode1s[i].Prejemnik.ToString());
            }

            #region OznakaRazpisa2Unfinished
            IWTable tableUnfinished2 = section.AddTable();
            tableUnfinished2.ResetCells(4, 5);
            tableUnfinished2.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableUnfinished2.TableFormat.BackColor = Color.White;
            tableUnfinished2.TableFormat.Paddings.All = 2;

            tableUnfinished2[0, 0].Width = 28f;
            tableUnfinished2[0, 1].Width = 30f;
            tableUnfinished2[0, 2].Width = 100f;
            WTextRange textRangeUnfinished2 = (WTextRange)tableUnfinished2[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeUnfinished2.CharacterFormat.Bold = true;
            textRangeUnfinished2.CharacterFormat.FontSize = 12f;
            tableUnfinished2[0, 3].Width = 100f;
            textRangeUnfinished2 = (WTextRange)tableUnfinished2[0, 3].AddParagraph().AppendText(_unfinishedTenderNumber2);
            textRangeUnfinished2.CharacterFormat.Bold = true;
            textRangeUnfinished2.CharacterFormat.FontSize = 12f;
            tableUnfinished2[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _unTenderCode2s.Count; i++)
            {
                tableUnfinished2[i + 1, 0].Width = 28f;
                tableUnfinished2[i + 1, 0].AddParagraph().AppendText(_unTenderCode2s[i].ZapStevilka1.ToString());
                tableUnfinished2[i + 1, 1].Width = 30f;
                tableUnfinished2[i + 1, 1].AddParagraph().AppendText(_unTenderCode2s[i].ZapStevilka2.ToString());
                tableUnfinished2[i + 1, 2].Width = 100f;
                tableUnfinished2[i + 1, 2].AddParagraph().AppendText(_unTenderCode2s[i].Oznaka1.ToString());
                tableUnfinished2[i + 1, 3].Width = 100f;
                tableUnfinished2[i + 1, 3].AddParagraph().AppendText(_unTenderCode2s[i].Oznaka2.ToString());
                tableUnfinished2[i + 1, 4].Width = 240f;
                tableUnfinished2[i + 1, 4].AddParagraph().AppendText(_unTenderCode2s[i].Prejemnik.ToString());
            }

            #region OznakaRazpisa3Unfinished
            IWTable tableUnfinished3 = section.AddTable();
            tableUnfinished3.ResetCells(4, 5);
            tableUnfinished3.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableUnfinished3.TableFormat.BackColor = Color.White;
            tableUnfinished3.TableFormat.Paddings.All = 2;

            tableUnfinished3[0, 0].Width = 28f;
            tableUnfinished3[0, 1].Width = 30f;
            tableUnfinished3[0, 2].Width = 100f;
            WTextRange textRangeUnfinished3 = (WTextRange)tableUnfinished3[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeUnfinished3.CharacterFormat.Bold = true;
            textRangeUnfinished3.CharacterFormat.FontSize = 12f;
            tableUnfinished3[0, 3].Width = 100f;
            textRangeUnfinished3 = (WTextRange)tableUnfinished3[0, 3].AddParagraph().AppendText(_unfinishedTenderNumber3);
            textRangeUnfinished3.CharacterFormat.Bold = true;
            textRangeUnfinished3.CharacterFormat.FontSize = 12f;
            tableUnfinished3[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _unTenderCode3s.Count; i++)
            {
                tableUnfinished3[i + 1, 0].Width = 28f;
                tableUnfinished3[i + 1, 0].AddParagraph().AppendText(_unTenderCode3s[i].ZapStevilka1.ToString());
                tableUnfinished3[i + 1, 1].Width = 30f;
                tableUnfinished3[i + 1, 1].AddParagraph().AppendText(_unTenderCode3s[i].ZapStevilka2.ToString());
                tableUnfinished3[i + 1, 2].Width = 100f;
                tableUnfinished3[i + 1, 2].AddParagraph().AppendText(_unTenderCode3s[i].Oznaka1.ToString());
                tableUnfinished3[i + 1, 3].Width = 100f;
                tableUnfinished3[i + 1, 3].AddParagraph().AppendText(_unTenderCode3s[i].Oznaka2.ToString());
                tableUnfinished3[i + 1, 4].Width = 240f;
                tableUnfinished3[i + 1, 4].AddParagraph().AppendText(_unTenderCode3s[i].Prejemnik.ToString());
            }

            SetTableHeadFinished(section);

            #region OznakaRazpisa1Finished
            IWTable tableFinished1 = section.AddTable();
            tableFinished1.ResetCells(4, 5);
            tableFinished1.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableFinished1.TableFormat.BackColor = Color.White;
            tableFinished1.TableFormat.Paddings.All = 2;

            tableFinished1[0, 0].Width = 28f;
            tableFinished1[0, 1].Width = 30f;
            tableFinished1[0, 2].Width = 100f;
            WTextRange textRangeFinished1 = (WTextRange)tableFinished1[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeFinished1.CharacterFormat.Bold = true;
            textRangeFinished1.CharacterFormat.FontSize = 12f;
            tableFinished1[0, 3].Width = 100f;
            textRangeFinished1 = (WTextRange)tableFinished1[0, 3].AddParagraph().AppendText(_finishedTenderNumber1);
            textRangeFinished1.CharacterFormat.Bold = true;
            textRangeFinished1.CharacterFormat.FontSize = 12f;
            tableFinished1[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _fiTenderCode1s.Count; i++)
            {
                tableFinished1[i + 1, 0].Width = 28f;
                tableFinished1[i + 1, 0].AddParagraph().AppendText(_fiTenderCode1s[i].ZapStevilka1.ToString());
                tableFinished1[i + 1, 1].Width = 30f;
                tableFinished1[i + 1, 1].AddParagraph().AppendText(_fiTenderCode1s[i].ZapStevilka2.ToString());
                tableFinished1[i + 1, 2].Width = 100f;
                tableFinished1[i + 1, 2].AddParagraph().AppendText(_fiTenderCode1s[i].Oznaka1.ToString());
                tableFinished1[i + 1, 3].Width = 100f;
                tableFinished1[i + 1, 3].AddParagraph().AppendText(_fiTenderCode1s[i].Oznaka2.ToString());
                tableFinished1[i + 1, 4].Width = 240f;
                tableFinished1[i + 1, 4].AddParagraph().AppendText(_fiTenderCode1s[i].Prejemnik.ToString());
            }

            #region OznakaRazpisa2Finished
            IWTable tableFinished2 = section.AddTable();
            tableFinished2.ResetCells(4, 5);
            tableFinished2.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableFinished2.TableFormat.BackColor = Color.White;
            tableFinished2.TableFormat.Paddings.All = 2;

            tableFinished2[0, 0].Width = 28f;
            tableFinished2[0, 1].Width = 30f;
            tableFinished2[0, 2].Width = 100f;
            WTextRange textRangeFinished2 = (WTextRange)tableFinished2[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeFinished2.CharacterFormat.Bold = true;
            textRangeFinished2.CharacterFormat.FontSize = 12f;
            tableFinished2[0, 3].Width = 100f;
            textRangeFinished2 = (WTextRange)tableFinished2[0, 3].AddParagraph().AppendText(_finishedTenderNumber1);
            textRangeFinished2.CharacterFormat.Bold = true;
            textRangeFinished2.CharacterFormat.FontSize = 12f;
            tableFinished2[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _fiTenderCode2s.Count; i++)
            {
                tableFinished2[i + 1, 0].Width = 28f;
                tableFinished2[i + 1, 0].AddParagraph().AppendText(_fiTenderCode2s[i].ZapStevilka1.ToString());
                tableFinished2[i + 1, 1].Width = 30f;
                tableFinished2[i + 1, 1].AddParagraph().AppendText(_fiTenderCode2s[i].ZapStevilka2.ToString());
                tableFinished2[i + 1, 2].Width = 100f;
                tableFinished2[i + 1, 2].AddParagraph().AppendText(_fiTenderCode2s[i].Oznaka1.ToString());
                tableFinished2[i + 1, 3].Width = 100f;
                tableFinished2[i + 1, 3].AddParagraph().AppendText(_fiTenderCode2s[i].Oznaka2.ToString());
                tableFinished2[i + 1, 4].Width = 240f;
                tableFinished2[i + 1, 4].AddParagraph().AppendText(_fiTenderCode2s[i].Prejemnik.ToString());
            }

            #region OznakaRazpisa3Finished
            IWTable tableFinished3 = section.AddTable();
            tableFinished3.ResetCells(4, 5);
            tableFinished3.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableFinished3.TableFormat.BackColor = Color.White;
            tableFinished3.TableFormat.Paddings.All = 2;

            tableFinished3[0, 0].Width = 28f;
            tableFinished3[0, 1].Width = 30f;
            tableFinished3[0, 2].Width = 100f;
            WTextRange textRangeFinished3 = (WTextRange)tableFinished3[0, 2].AddParagraph().AppendText(_tenderUnit);
            textRangeFinished3.CharacterFormat.Bold = true;
            textRangeFinished3.CharacterFormat.FontSize = 12f;
            tableFinished3[0, 3].Width = 100f;
            textRangeFinished3 = (WTextRange)tableFinished3[0, 3].AddParagraph().AppendText(_finishedTenderNumber3);
            textRangeFinished3.CharacterFormat.Bold = true;
            textRangeFinished3.CharacterFormat.FontSize = 12f;
            tableFinished3[0, 4].Width = 240f;
            #endregion
            for (int i = 0; i < _fiTenderCode3s.Count; i++)
            {
                tableFinished3[i + 1, 0].Width = 28f;
                tableFinished3[i + 1, 0].AddParagraph().AppendText(_fiTenderCode3s[i].ZapStevilka1.ToString());
                tableFinished3[i + 1, 1].Width = 30f;
                tableFinished3[i + 1, 1].AddParagraph().AppendText(_fiTenderCode3s[i].ZapStevilka2.ToString());
                tableFinished3[i + 1, 2].Width = 100f;
                tableFinished3[i + 1, 2].AddParagraph().AppendText(_fiTenderCode3s[i].Oznaka1.ToString());
                tableFinished3[i + 1, 3].Width = 100f;
                tableFinished3[i + 1, 3].AddParagraph().AppendText(_fiTenderCode3s[i].Oznaka2.ToString());
                tableFinished3[i + 1, 4].Width = 240f;
                tableFinished3[i + 1, 4].AddParagraph().AppendText(_fiTenderCode3s[i].Prejemnik.ToString());
            }


            #region Saving document 
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }

        private IWParagraph SetTableHeadUnfinished(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            IWTable tableHeadUnfinished = section.AddTable();
            tableHeadUnfinished.ResetCells(1, 5);
            tableHeadUnfinished.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableHeadUnfinished.TableFormat.BackColor = Color.White;
            tableHeadUnfinished.TableFormat.Paddings.All = 2;

            tableHeadUnfinished[0, 0].Width = 28f;
            tableHeadUnfinished[0, 1].Width = 30f;
            tableHeadUnfinished[0, 2].Width = 100f;
            WTextRange textRangeUnfinished = (WTextRange)tableHeadUnfinished[0, 2].AddParagraph().AppendText(_investStatusText);
            textRangeUnfinished.CharacterFormat.Bold = true;
            textRangeUnfinished.CharacterFormat.FontSize = 14f;
            textRangeUnfinished.CharacterFormat.TextBackgroundColor = Color.Yellow;
            tableHeadUnfinished[0, 3].Width = 100f;
            textRangeUnfinished = (WTextRange)tableHeadUnfinished[0, 3].AddParagraph().AppendText(_unfinishedText);
            textRangeUnfinished.CharacterFormat.Bold = true;
            textRangeUnfinished.CharacterFormat.FontSize = 14f;
            textRangeUnfinished.CharacterFormat.TextBackgroundColor = Color.Yellow;
            tableHeadUnfinished[0, 4].Width = 240f;
            return paragraph;
        }

        private void SetTableHeadFinished(IWSection section)
        {
            IWTable tableHeadFinished = section.AddTable();
            tableHeadFinished.ResetCells(1, 5);
            tableHeadFinished.TableFormat.HorizontalAlignment = RowAlignment.Left;
            tableHeadFinished.TableFormat.BackColor = Color.White;
            tableHeadFinished.TableFormat.Paddings.All = 2;

            tableHeadFinished[0, 0].Width = 28f;
            tableHeadFinished[0, 1].Width = 30f;
            tableHeadFinished[0, 2].Width = 100f;
            WTextRange textRangeFinished = (WTextRange)tableHeadFinished[0, 2].AddParagraph().AppendText(_investStatusText);
            textRangeFinished.CharacterFormat.Bold = true;
            textRangeFinished.CharacterFormat.FontSize = 14f;
            textRangeFinished.CharacterFormat.TextBackgroundColor = Color.Yellow;
            tableHeadFinished[0, 3].Width = 100f;
            textRangeFinished = (WTextRange)tableHeadFinished[0, 3].AddParagraph().AppendText(_finishedText);
            textRangeFinished.CharacterFormat.Bold = true;
            textRangeFinished.CharacterFormat.FontSize = 14f;
            textRangeFinished.CharacterFormat.TextBackgroundColor = Color.Yellow;
            tableHeadFinished[0, 4].Width = 240f;
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
        private IWParagraph SetDateRemittances(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRange2 = paragraph.AppendText(_dateRemittances) as WTextRange;
            textRange2.CharacterFormat.FontName = "Calibri";
            textRange2.CharacterFormat.FontSize = 12f;
            textRange2.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }
        private IWParagraph SetHeadDocument(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange1 = paragraph.AppendText(_head) as WTextRange;
            textRange1.CharacterFormat.FontName = "Calibri";
            textRange1.CharacterFormat.FontSize = 16f;
            textRange1.CharacterFormat.TextColor = Color.Black;
            textRange1.CharacterFormat.Bold = true;
            return paragraph;
        }
    }
}
