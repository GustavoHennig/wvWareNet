using System.Collections.Generic;

namespace WvWareNet.Core
{
    public class DocumentModel
    {
        public FileInformationBlock FileInfo { get; set; }
        public List<Section> Sections { get; } = new List<Section>();
        public List<HeaderFooter> Headers { get; } = new List<HeaderFooter>();
        public List<HeaderFooter> Footers { get; } = new List<HeaderFooter>();
        public List<Footnote> Footnotes { get; } = new List<Footnote>();
        public List<TextBox> TextBoxes { get; } = new List<TextBox>();
    }

    public class TextBox
    {
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();
    }

    public class Section
    {
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();
        public SectionProperties Properties { get; set; }
    }

    public class Paragraph
    {
        public List<Run> Runs { get; } = new List<Run>();
        public ParagraphProperties Properties { get; set; }
    }

    public class Run
    {
        public string Text { get; set; }
        public CharacterProperties Properties { get; set; }
    }

    public class Footnote
    {
        public int ReferenceId { get; set; }
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();
    }

    public class HeaderFooter
    {
        public HeaderFooterType Type { get; set; }
        public List<Paragraph> Paragraphs { get; } = new List<Paragraph>();
    }

    public enum HeaderFooterType
    {
        FirstPage,
        EvenPage,
        OddPage,
        Default
    }

    // Existing properties classes are already implemented
    // (FileInformationBlock, CharacterProperties, ParagraphProperties, SectionProperties)
}
