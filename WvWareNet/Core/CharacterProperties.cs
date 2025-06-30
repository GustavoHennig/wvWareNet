namespace WvWareNet.Core;

public class CharacterProperties
{
    // Font codes
    public ushort FtcaAscii { get; set; }
    public ushort FtcaFE { get; set; }
    public ushort FtcaOther { get; set; }
    
    // Font size in half-points
    public ushort FontSize { get; set; }
    
    // Language identifiers
    public ushort LanguageId { get; set; }
    public ushort DefaultLanguageId { get; set; }
    public ushort FarEastLanguageId { get; set; }
    
    // Formatting flags
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderlined { get; set; }
    public bool IsStrikeThrough { get; set; }
    public bool IsSmallCaps { get; set; }
    public bool IsAllCaps { get; set; }
    public bool IsHidden { get; set; }
    
    // Character spacing
    public short CharacterSpacing { get; set; }
    
    // Character scaling (percentage)
    public ushort CharacterScaling { get; set; }
    
    public CharacterProperties()
    {
        // Set default values
        FtcaAscii = 0;
        FtcaFE = 0;
        FtcaOther = 0;
        FontSize = 20; // Default font size (20 half-points = 10pt)
        LanguageId = 0;
        DefaultLanguageId = 0x0400; // English (US)
        FarEastLanguageId = 0x0400;
        CharacterScaling = 100; // 100% scaling
    }

    public object Clone()
    {
        return this.MemberwiseClone();
    }
}
