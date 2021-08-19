using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableMergeColumns_Word
{
    public class DocumentGenerator
    {
        #region PROPS
        string _measureMainName;
        string _scopeMeasureMainName;
        string _recognizedCostsMainName;
        string _amountIncentiveMainName;
        string _incentiveSumName;
        decimal _incentiveSumData;
        List<TableData> _tabledatas;
        #endregion

        #region CTOR
        public DocumentGenerator(
        string aMeasureMainName,
        string aScopeMeasureMainName,
        string aRecognizedCostsMainName,
        string aAmountIncentiveMainName,
        string aIncentiveSumName,
        decimal aIncentiveSumData,
        List<TableData> aTableDatas)
        {
            _measureMainName = aMeasureMainName;
            _scopeMeasureMainName = aScopeMeasureMainName;
            _recognizedCostsMainName = aRecognizedCostsMainName;
            _amountIncentiveMainName = aAmountIncentiveMainName;
            _incentiveSumName = aIncentiveSumName;
            _incentiveSumData = aIncentiveSumData;
            _tabledatas = aTableDatas;
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
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            IWTable table = section.AddTable();
            table.ResetCells(1, 5);
            table.TableFormat.Paddings.All = 2;
            table.TableFormat.Borders.Horizontal.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Vertical.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Left.BorderType = BorderStyle.None;
            table.TableFormat.Borders.Right.BorderType = BorderStyle.None;

            table[0, 0].Width = 20f;

            table[0, 1].AddParagraph().AppendText(_measureMainName);
            table[0, 1].Width = 210f;

            table[0, 2].AddParagraph().AppendText(_scopeMeasureMainName);
            table[0, 2].Width = 80f;

            table[0, 3].AddParagraph().AppendText(_recognizedCostsMainName);
            table[0, 3].Width = 80f;

            table[0, 4].AddParagraph().AppendText(_amountIncentiveMainName);
            table[0, 4].Width = 80f;

            int i = 1;
            foreach (var plannedInvestment in _tabledatas)
            {
                WTableRow tableRow = table.AddRow(true);
                tableRow.RowFormat.Borders.Vertical.BorderType = BorderStyle.None;
                tableRow.RowFormat.Borders.Horizontal.BorderType = BorderStyle.None;

                tableRow.Cells[0].AddParagraph().AppendText(i.ToString());
                tableRow.Cells[0].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
                tableRow.Cells[1].AddParagraph().AppendText(plannedInvestment.Measure);
                tableRow.Cells[1].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
                tableRow.Cells[2].AddParagraph().AppendText(plannedInvestment.ScopeMeasure);
                tableRow.Cells[2].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
                tableRow.Cells[3].AddParagraph().AppendText(Convert.ToDecimal(plannedInvestment.RecognizedCosts).ToString("C"));
                tableRow.Cells[3].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
                tableRow.Cells[4].AddParagraph().AppendText(plannedInvestment.AmountIncentive.ToString("N", new CultureInfo("is-IS")));
                tableRow.Cells[4].CellFormat.Borders.Top.BorderType = BorderStyle.Single;

                WTableRow tableRow2 = table.AddRow(true);

                tableRow2.Cells[0].AddParagraph().AppendText("");
                tableRow2.Cells[0].CellFormat.Borders.Top.BorderType = BorderStyle.None;
                tableRow2.Cells[1].AddParagraph().AppendText(plannedInvestment.MergeDescriptionAdds);
                tableRow2.Cells[1].CellFormat.Borders.Top.BorderType = BorderStyle.None;
                i++;
            }

            for (int row = 1; row <= _tabledatas.Count(); row++)
            {
                //Console.WriteLine(row * 2);

                table.ApplyHorizontalMerge(row * 2, 1, 4);
            }

            //WTableRow tableRowSum = tableMergeText.AddRow(true);

            //tableRowSum.Cells[3].AddParagraph().AppendText(_incentiveSumName);

            ////var totalIWTextRage = cell.AddParagraph().AppendText(_incentiveSumName);
            ////totalIWTextRage.CharacterFormat.Bold = true;

            //tableRowSum.Cells[4].AddParagraph().AppendText(_tabledatas.Sum(p => p.AmountIncentive).ToString("N", new CultureInfo("is-IS")));
            ////totalAmountIWTextRage.CharacterFormat.Bold = true;

            #region TableSum = Spodbuda skupaj
            IWTable tableRowSum = section.AddTable();
            tableRowSum.ResetCells(1, 5);
            tableRowSum.TableFormat.Borders.Bottom.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Vertical.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Horizontal.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Left.BorderType = BorderStyle.None;
            tableRowSum.TableFormat.Borders.Right.BorderType = BorderStyle.None;
            tableRowSum[0, 0].AddParagraph().AppendText("");
            tableRowSum[0, 0].Width = 20f;
            tableRowSum[0, 0].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            tableRowSum[0, 1].AddParagraph().AppendText("");
            tableRowSum[0, 1].Width = 210f;
            tableRowSum[0, 1].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            tableRowSum[0, 2].AddParagraph().AppendText("");
            tableRowSum[0, 2].Width = 70f;
            tableRowSum[0, 2].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            var wtextRangeTableRowSum = tableRowSum[0, 3].AddParagraph().AppendText(_incentiveSumName);
            tableRowSum[0, 3].Width = 90f;
            tableRowSum[0, 3].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            wtextRangeTableRowSum.CharacterFormat.Bold = true;
            wtextRangeTableRowSum = tableRowSum[0, 4].AddParagraph().AppendText(_tabledatas.Sum(p => p.AmountIncentive).ToString("N", new CultureInfo("is-IS")));
            tableRowSum[0, 4].Width = 80f;
            tableRowSum[0, 4].CellFormat.Borders.Top.BorderType = BorderStyle.Single;
            wtextRangeTableRowSum.CharacterFormat.Bold = true;
            #endregion

            #region MemoryStream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }
    }
}
