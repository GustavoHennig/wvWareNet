namespace WvWareNet.Core
{
    /// <summary>
    /// Represents section properties in a Word document
    /// </summary>
    public class SectionProperties
    {
        // Page size and margins
        public int PageWidth { get; set; } = 12240;  // Default: 8.5" in twips (1/1440 inch)
        public int PageHeight { get; set; } = 15840; // Default: 11" in twips
        public int LeftMargin { get; set; } = 1800;  // Default: 1.25" in twips
        public int RightMargin { get; set; } = 1800; // Default: 1.25" in twips
        public int TopMargin { get; set; } = 1440;   // Default: 1" in twips
        public int BottomMargin { get; set; } = 1440; // Default: 1" in twips

        // Page orientation
        public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

        // Column properties
        public int ColumnCount { get; set; } = 1;
        public int ColumnSpacing { get; set; } = 720; // Default: 0.5" in twips

        // Page numbering
        public PageNumberFormat PageNumberFormat { get; set; } = PageNumberFormat.Arabic;
        public int StartingPageNumber { get; set; } = 1;

        // Header/footer distances
        public int HeaderDistance { get; set; } = 720;  // Default: 0.5" in twips
        public int FooterDistance { get; set; } = 720;  // Default: 0.5" in twips
    }

    public enum PageOrientation
    {
        Portrait,
        Landscape
    }

    public enum PageNumberFormat
    {
        Arabic,
        RomanUpper,
        RomanLower,
        LetterUpper,
        LetterLower
    }
}
