using System.Collections.Generic;

namespace WvWareNet.Core
{
    public class DocumentModel
    {
        public List<Paragraph> Paragraphs { get; set; } = new();
    }

    public class Paragraph
    {
        public List<Run> Runs { get; set; } = new();
    }

    public class Run
    {
        public string Text { get; set; }
        // Add formatting properties as needed
    }
}
