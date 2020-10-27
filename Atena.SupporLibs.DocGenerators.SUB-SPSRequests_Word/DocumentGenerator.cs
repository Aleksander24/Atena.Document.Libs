using System;
using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Linq;
using Atena.SupporLibs.DocGenerators.SUB_SPSRequests_Word.Models;
using System.Collections.Generic;
using System.Globalization;
using Image = Syncfusion.Drawing.Image;

namespace Atena.SupporLibs.DocGenerators.SUB_SPSRequests_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";
        public string Label => "Test";
        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        #region PROPS
        string _sender;
        string _receiver;
        string _transferRequest;
        string _transferRequestCont;
        string _date;
        string _publicTenderText;
        string _programFunds;
        List<MainTableRowsData> _rowDatas;
        List<SPSRecapitulationData> _spsRecapitulations;
        string _prepared;
        string _responsiblePerson;
        string _attachments;
        #endregion
        #region DocumentGenerator
        public DocumentGenerator(
            string aSender, 
            string aReceiver, 
            string aTransferRequest, 
            string aTransferRequestCont, 
            string aDate, 
            string aPublicTenderText, 
            string aProgramFunds, 
            List<MainTableRowsData> aRowDatas, 
            List<SPSRecapitulationData> aSPSRecapitulations, 
            string aPrepared, 
            string aResponsiblePerson, 
            string aAttachments)

        {
            _sender = aSender;
            _receiver = aReceiver;
            _transferRequest = aTransferRequest;
            _transferRequestCont = aTransferRequestCont;
            _date = aDate;
            _publicTenderText = aPublicTenderText;
            _programFunds = aProgramFunds;
            _rowDatas = aRowDatas;
            _spsRecapitulations = aSPSRecapitulations;
            _prepared = aPrepared;
            _responsiblePerson = aResponsiblePerson;
            _attachments = aAttachments;
        }
        #endregion

        public byte[] Generate()
        {
            #region Creating document, add section, set Margin, create paragraph
            //Creating a new document
            WordDocument document = new WordDocument();

            // Adding a nes section to the document
            IWSection section = document.AddSection();

            // Set Margin of the section
            section.PageSetup.Margins.All = 40;

            // Set page size of the section
            section.PageSetup.PageSize = new SizeF(575, 792);

            // Create Paragraph styles for normal font
            WParagraphStyle style = document.AddParagraphStyle("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Calibri";
            style.CharacterFormat.FontSize = 9f;
            style.ParagraphFormat.BeforeSpacing = 0;
            style.ParagraphFormat.AfterSpacing = 0;
            style.ParagraphFormat.LineSpacing = 10f;
            style.CharacterFormat.TextColor = Color.Black;

            // Create paragraph
            IWParagraph paragraph = section.HeadersFooters.Header.AddParagraph();
            #endregion

            paragraph = SetSender(_sender, section);

            paragraph = SetReceiver(_receiver, section);

            paragraph = SetTransferRequest(_transferRequest, section);
            SetTransferRequestCont(_transferRequestCont, paragraph);
            SetDate(_date, paragraph);

            paragraph = SetPublicTenderName(_publicTenderText, section);

            paragraph = SetProgramFunds(_programFunds, section);


            #region Creating Main table
            // Creating Main table
            IWTable table = section.AddTable();
            table.ResetCells(6, 8);
            table.TableFormat.BackColor = Color.White;
            table.TableFormat.HorizontalAlignment = RowAlignment.Left;
            table.TableFormat.Positioning.HorizPosition = 1;
            table.TableFormat.Positioning.VertPosition = 230;
            table.TableFormat.Paddings.All = 2;
            #endregion

            #region Headings for Main table
            table[0, 0].Width = 28f;
            table[0, 0].AddParagraph().AppendText("Zap št.");
            table[0, 1].Width = 30f;
            table[0, 1].AddParagraph().AppendText("");
            table[0, 2].Width = 80f;
            table[0, 2].AddParagraph().AppendText("Številka pogodbe");
            table[0, 3].Width = 85f;
            table[0, 3].AddParagraph().AppendText("Prejemnik");
            table[0, 4].Width = 85f;
            table[0, 4].AddParagraph().AppendText("Naslov");
            table[0, 5].Width = 76f;
            table[0, 5].AddParagraph().AppendText("Pošta");
            table[0, 6].Width = 54.4f;
            table[0, 6].AddParagraph().AppendText("Davčna številka");
            table[0, 7].Width = 60f;
            table[0, 7].AddParagraph().AppendText("Vrednost v EUR");
            #endregion

            // rows in Main table
            for (int i = 0; i < _rowDatas.Count; i++)
            {
                table[i + 1, 0].Width = 28f;
                table[i + 1, 0].AddParagraph().AppendText(_rowDatas[i].ZapStevilka.ToString());
                table[i + 1, 1].Width = 30f;
                table[i + 1, 1].AddParagraph().AppendText(_rowDatas[i].RegularStevilka.ToString());
                table[i + 1, 2].Width = 80f;
                table[i + 1, 2].AddParagraph().AppendText(_rowDatas[i].StevilkaPogodbe.ToString());
                table[i + 1, 3].Width = 85f;
                table[i + 1, 3].AddParagraph().AppendText(_rowDatas[i].Prejemnik.ToString());
                table[i + 1, 4].Width = 85f;
                table[i + 1, 4].AddParagraph().AppendText(_rowDatas[i].Naslov.ToString());
                table[i + 1, 5].Width = 76f;
                table[i + 1, 5].AddParagraph().AppendText(_rowDatas[i].Posta.ToString());
                table[i + 1, 6].Width = 54.4f;
                table[i + 1, 6].AddParagraph().AppendText(_rowDatas[i].DavcnaStevilka.ToString());
                table[i + 1, 7].Width = 60f;
                table[i + 1, 7].AddParagraph().AppendText(_rowDatas[i].VrednostVEUR.ToString());

            }

            #region Skupaj at the end of the Main table (for "Vrednost v EUR)
            IWTable table1 = section.AddTable();
            table1.ResetCells(1, 2);
            table1.TableFormat.BackColor = Color.White;
            table1.TableFormat.HorizontalAlignment = RowAlignment.Right;
            table1.TableFormat.Paddings.All = 2;
            table1[0, 0].Width = 54.5f;
            table1[0, 0].AddParagraph().AppendText("Skupaj:");
            // sum for table
            decimal sumOfAllRowDatas = _rowDatas.Sum(p => p.VrednostVEUR);
            table1[0, 1].Width = 60f;
            table1[0, 1].AddParagraph().AppendText(sumOfAllRowDatas.ToString());
            #endregion

            paragraph = SetPrepared(_prepared, section);

            paragraph = SetResponsiblePerson(_responsiblePerson, section);

            paragraph = SetAttachments(_attachments, section);

            paragraph = ImageSignatureResponsiblePerson(section);

            paragraph = ImageSignaturePrepared(section);

            #region Adding for naslov Rekapitulacija
            WParagraphStyle styleBox = document.AddParagraphStyle("Naslov Rekapitulacija") as WParagraphStyle;
            styleBox.CharacterFormat.FontName = "Calibri";
            styleBox.CharacterFormat.FontSize = 14f;
            styleBox.ParagraphFormat.BeforeSpacing = 0;
            styleBox.ParagraphFormat.AfterSpacing = 0;
            styleBox.ParagraphFormat.LineSpacing = 15f;
            #endregion

            #region Adding textbox and Rekapitulacija heading
            paragraph = section.AddParagraph();
            IWTextBox textBox = paragraph.AppendTextBox(260, 95);
            textBox.TextBoxFormat.HorizontalAlignment = ShapeHorizontalAlignment.Center;
            textBox.TextBoxFormat.FillColor = Color.LightGray;
            WParagraph textBoxParagraph = textBox.TextBoxBody.AddParagraph() as WParagraph;
            textBoxParagraph.ApplyStyle("Naslov Rekapitulacija");
            textBoxParagraph.AppendText("Rekapitulacija zahtevka po projektih");
            textBoxParagraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            #endregion

            #region Table inside SPSRecapitulation
            WTable tableRec = textBox.TextBoxBody.AddTable() as WTable;
            tableRec.ResetCells(3, 2);
            tableRec.TableFormat.BackColor = Color.White;
            tableRec.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableRec.TableFormat.Paddings.All = 2;

            //Head of table SPSRecapitulation
            string SPSProject = "SPS projekt: ";
            string SumProject = "Vsota projekta: ";
            tableRec[0, 0].Width = 185f;
            tableRec[0, 0].AddParagraph().AppendText(SPSProject);
            tableRec[0, 1].Width = 73f;
            tableRec[0, 1].AddParagraph().AppendText(SumProject);
            #endregion

            // rows in table SPSRecapitulation
            for (int i = 0; i < _spsRecapitulations.Count; i++)
            {
                tableRec[i + 1, 0].Width = 185f;
                tableRec[i + 1, 0].AddParagraph().AppendText(_spsRecapitulations[i].SPSProjectName.ToString());
                tableRec[i + 1, 1].Width = 73f;
                tableRec[i + 1, 1].AddParagraph().AppendText(_spsRecapitulations[i].SPSProjectSum.ToString());
            }

            #region End of table SPSRecapitulation (Vsota zahtevka)
            string SumRequest = "Vsota zahtevka: ";
            WTable tableRecEnd = textBox.TextBoxBody.AddTable() as WTable;
            tableRecEnd.ResetCells(1, 2);
            tableRecEnd.TableFormat.BackColor = Color.White;
            tableRecEnd.TableFormat.HorizontalAlignment = RowAlignment.Center;
            tableRecEnd.TableFormat.Paddings.All = 2;
            tableRecEnd[0, 0].Width = 185f;
            tableRecEnd[0, 0].AddParagraph().AppendText(SumRequest);
            //sum SPSRecapitulation
            decimal sumOfAllSPSProjects = _spsRecapitulations.Sum(p => p.SPSProjectSum);
            tableRecEnd[0, 1].Width = 73f;
            tableRecEnd[0, 1].AddParagraph().AppendText(sumOfAllSPSProjects.ToString());
            #endregion

            #region Saving word document
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }
        private static IWParagraph ImageSignaturePrepared(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            FileStream imageStream = new FileStream($"J:\\PROJEKTI\\ATENA_SUPPORT\\Atena.Document.Libs\\Atena.SupporLibs.DocGenerators.SUB-SPSRequests_Word\\Images\\Uefa_logo.png", FileMode.Open, FileAccess.Read);
            IWPicture picture = paragraph.AppendPicture(imageStream);
            picture.TextWrappingStyle = TextWrappingStyle.Square;
            picture.Width = 30;
            picture.Height = 30;
            picture.VerticalPosition = 345; // set up static
            picture.HorizontalAlignment = ShapeHorizontalAlignment.Left;
            return paragraph;
        }
        private static IWParagraph ImageSignatureResponsiblePerson(IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            FileStream imageStream1 = new FileStream($"J:\\PROJEKTI\\ATENA_SUPPORT\\Atena.Document.Libs\\Atena.SupporLibs.DocGenerators.SUB-SPSRequests_Word\\Images\\EA_sports.png", FileMode.Open, FileAccess.Read);
            IWPicture picture1 = paragraph.AppendPicture(imageStream1);
            picture1.TextWrappingStyle = TextWrappingStyle.Square;
            picture1.Width = 30;
            picture1.Height = 30;
            picture1.VerticalPosition = 370; // set up static
                                                //picture.HorizontalPosition = 1;
            picture1.HorizontalAlignment = ShapeHorizontalAlignment.Right;
            //picture.VerticalOrigin = VerticalOrigin.Margin;
            //picture.VerticalPosition = 100;
            //picture.HorizontalOrigin = HorizontalOrigin.RightMargin;
            //picture.HorizontalPosition = 500f;
            //picture.WidthScale = 20;
            //picture.HeightScale = 15;
            return paragraph;
        }
        private static IWParagraph SetAttachments(string aAttachments, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange10 = paragraph.AppendText(aAttachments) as WTextRange;
            textRange10.CharacterFormat.FontSize = 9f;
            textRange10.CharacterFormat.FontName = "Calibri";
            textRange10.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }
        private static IWParagraph SetResponsiblePerson(string aResponsiblePerson, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Right;
            WTextRange textRange9 = paragraph.AppendText(aResponsiblePerson) as WTextRange;
            textRange9.CharacterFormat.FontSize = 9f;
            textRange9.CharacterFormat.FontName = "Calibri";
            textRange9.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            return paragraph;
        }
        private static IWParagraph SetPrepared(string aPrepared, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange8 = paragraph.AppendText(aPrepared) as WTextRange;
            textRange8.CharacterFormat.FontSize = 9f;
            textRange8.CharacterFormat.FontName = "Calibri";
            textRange8.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            return paragraph;
        }
        private static IWParagraph SetProgramFunds(string aProgramFunds, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange7 = paragraph.AppendText(aProgramFunds) as WTextRange;
            textRange7.CharacterFormat.FontSize = 10.25f;
            textRange7.CharacterFormat.FontName = "Calibri";
            textRange7.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            return paragraph;
        }
        private static IWParagraph SetPublicTenderName(string aPublicTenderText, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange6 = paragraph.AppendText(aPublicTenderText) as WTextRange;
            textRange6.CharacterFormat.FontSize = 10f;
            textRange6.CharacterFormat.FontName = "Calibri";
            textRange6.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }
        private static void SetDate(string aDate, IWParagraph paragraph)
        {
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange5 = paragraph.AppendText(aDate) as WTextRange;
            textRange5.CharacterFormat.FontSize = 9f;
            textRange5.CharacterFormat.FontName = "Calibri";
            textRange5.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
        }
        private static void SetTransferRequestCont(string aTransferRequestCont, IWParagraph paragraph)
        {
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange4 = paragraph.AppendText(aTransferRequestCont) as WTextRange;
            textRange4.CharacterFormat.FontSize = 14f;
            textRange4.CharacterFormat.FontName = "Calibri";
            textRange4.CharacterFormat.Bold = true;
        }
        private static IWParagraph SetTransferRequest(string aTransferRequest, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange3 = paragraph.AppendText(aTransferRequest) as WTextRange;
            textRange3.CharacterFormat.FontSize = 14f;
            textRange3.CharacterFormat.FontName = "Calibri";
            textRange3.CharacterFormat.TextColor = Color.Black;
            return paragraph;
        }
        private static IWParagraph SetReceiver(string aReceiver, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange2 = paragraph.AppendText(aReceiver) as WTextRange;
            textRange2.CharacterFormat.FontSize = 10f;
            textRange2.CharacterFormat.FontName = "Calibri";
            textRange2.CharacterFormat.TextColor = Syncfusion.Drawing.Color.Black;
            textRange2.CharacterFormat.Bold = true;
            return paragraph;
        }
        private static IWParagraph SetSender(string _sender, IWSection section)
        {
            IWParagraph paragraph = section.AddParagraph();
            paragraph.ApplyStyle("Normal");
            paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Left;
            WTextRange textRange1 = paragraph.AppendText(_sender) as WTextRange;
            textRange1.CharacterFormat.FontSize = 10f;
            textRange1.CharacterFormat.FontName = "Calibri";
            textRange1.CharacterFormat.TextColor = Color.Black;
            textRange1.CharacterFormat.Bold = true;
            return paragraph;
        }
    }
}
