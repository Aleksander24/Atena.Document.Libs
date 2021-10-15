using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TableMergeColumns2_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            

            var time1 = DateTime.Now.ToFileTime().ToString();
            File.WriteAllBytes($"C:\\Users\\aleks\\Desktop\\DeloOdDoma\\test\\TableMergeColumnUpravičenci{time1}.docx", Generate());

        }

        public static byte[] Generate()
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
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            IWTable table = section.AddTable();
            table.ResetCells(1, 4);
            table.TableFormat.Paddings.All = 2;
            table.TableFormat.Borders.Horizontal.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Vertical.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Left.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Right.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Top.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Bottom.BorderType = BorderStyle.None;

            var imenaUpravičencev = new List<TableData>()
            {
                new TableData()
                {
                    imenaUpravičencev = "Janez",
                    delež = 2.00M,
                    spodbuda = 33.0M
                },
                new TableData()
                {
                    imenaUpravičencev = "Janko",
                    delež = 2.00M,
                    spodbuda = 33.0M
                },
                new TableData()
                {
                    imenaUpravičencev = "Metka",
                    delež = 2.00M,
                    spodbuda = 33.0M
                }
            };
            int i = 1;
            foreach (var imenaUpravičenca in imenaUpravičencev)
            {
                WTableRow tableRow1 = table.AddRow(true);
                tableRow1.RowFormat.Borders.Vertical.BorderType = BorderStyle.None;
                tableRow1.RowFormat.Borders.Horizontal.BorderType = BorderStyle.None;
                tableRow1.RowFormat.Borders.Top.BorderType = BorderStyle.None;
                tableRow1.RowFormat.Borders.Bottom.BorderType = BorderStyle.None;

                

                var textRange = tableRow1.Cells[0].AddParagraph().AppendText(i.ToString());
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 9;
                tableRow1.Cells[0].Width = 50f;
                //tableRow1.Cells[0].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

                textRange = tableRow1.Cells[1].AddParagraph().AppendText(imenaUpravičenca.imenaUpravičencev);
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 9;
                tableRow1.Cells[1].Width = 280f;
                //tableRow1.Cells[1].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

                textRange = tableRow1.Cells[2].AddParagraph().AppendText(imenaUpravičenca.delež.ToString());
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 9;
                tableRow1.Cells[2].Width = 80f;
                //tableRow1.Cells[2].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

                textRange = tableRow1.Cells[3].AddParagraph().AppendText(imenaUpravičenca.spodbuda.ToString());
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 9;
                tableRow1.Cells[3].Width = 80f;
                //tableRow1.Cells[3].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
                i++;
            }
            var textRange1 = table[0, 0].AddParagraph().AppendText("Številka");
            table[0, 0].Width = 50f;
            table[0, 0].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            //tableRow.Cells[0].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;

            textRange1 = table[0, 1].AddParagraph().AppendText("Upravičena oseba");
            table[0, 1].Width = 280f;
            table[0, 1].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            //tableRow.Cells[1].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;

            textRange1 = table[0, 2].AddParagraph().AppendText("delež [%]");
            table[0, 2].Width = 80f;
            table[0, 2].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            //tableRow.Cells[2].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;

            textRange1 = table[0, 3].AddParagraph().AppendText("finančne spodbude [EUR]");
            table[0, 3].Width = 80f;
            table[0, 3].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            //tableRow.Cells[3].CellFormat.Borders.Bottom.BorderType = BorderStyle.Single;

            IWTable tableRowSum = section.AddTable();
            tableRowSum.ResetCells(1, 3);
            tableRowSum.TableFormat.Borders.Bottom.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Vertical.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Horizontal.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Left.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Right.BorderType = BorderStyle.None;

            textRange1 = tableRowSum[0, 0].AddParagraph().AppendText("VIŠINA NEPOVRATNE FINANČNE SPODBUDE");
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            textRange1.CharacterFormat.Bold = true;
            tableRowSum[0, 0].Width = 330f;
            tableRowSum[0, 0].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

            textRange1 = tableRowSum[0, 1].AddParagraph().AppendText(imenaUpravičencev.Sum(p => p.delež).ToString());
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            textRange1.CharacterFormat.Bold = true;
            tableRowSum[0, 1].Width = 80f;
            tableRowSum[0, 1].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

            textRange1 = tableRowSum[0, 2].AddParagraph().AppendText(imenaUpravičencev.Sum(p => p.spodbuda).ToString());
            textRange1.CharacterFormat.FontName = "Arial";
            textRange1.CharacterFormat.FontSize = 9;
            textRange1.CharacterFormat.Bold = true;
            tableRowSum[0, 2].Width = 80f;
            tableRowSum[0, 2].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            


            #region MemoryStream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }


        public class TableData
        {
            
            public string imenaUpravičencev { get; set; }
            public decimal delež { get; set; }
            public decimal spodbuda { get; set; }
        }
    }
}
