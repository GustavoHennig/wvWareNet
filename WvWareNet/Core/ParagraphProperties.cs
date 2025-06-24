namespace WvWareNet.Core;

public class ParagraphProperties
{
    // Justification
    public Justification Justification { get; set; }
    
    // Indentation
    public int LeftIndent { get; set; }
    public int RightIndent { get; set; }
    public int FirstLineIndent { get; set; }
    
    // Spacing
    public int SpaceBefore { get; set; }
    public int SpaceAfter { get; set; }
    public int LineSpacing { get; set; }
    
    // Borders and shading
    public bool HasBorder { get; set; }
    public bool HasShading { get; set; }
    
    // Tab settings
    public TabStop[] TabStops { get; set; }
    
    // Page break control
    public bool PageBreakBefore { get; set; }
    public bool KeepWithNext { get; set; }
    public bool KeepTogether { get; set; }
    
    // Style information
    public int StyleIndex { get; set; }
    
    public ParagraphProperties()
    {
        Justification = Justification.Left;
        TabStops = Array.Empty<TabStop>();
    }
}

public enum Justification
{
    Left,
    Center,
    Right,
    Justified
}

public struct TabStop
{
    public int Position { get; set; }
    public TabAlignment Alignment { get; set; }
    public TabLeader Leader { get; set; }
}

public enum TabAlignment
{
    Left,
    Center,
    Right,
    Decimal,
    Bar
}

public enum TabLeader
{
    None,
    Dotted,
    Dashed,
    Underline,
    ThickLine,
    DoubleLine
}
