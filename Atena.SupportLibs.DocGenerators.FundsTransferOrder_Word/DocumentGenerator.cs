using Atena.SupportLibs.Core.Enum;
using Atena.SupportLibs.Core.Interfaces;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.Drawing;

namespace Atena.SupportLibs.DocGenerators.FundsTransferOrder_Word
{
    public class DocumentGenerator : IDocumentGenerator
    {
        public string Version => "1.0.0";

        public string Label => "DemoTest_FundsTransferOrder";

        public DocumentTypeEnum DocumentTypeEnum => DocumentTypeEnum.Word;

        public DocumentGenerator()
        {

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

            #region Saving document to stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream, FormatType.Docx);
            stream.Position = 0;
            return stream.ToArray();
            #endregion
        }
    }
}
