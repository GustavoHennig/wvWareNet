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
    /// <summary>
    /// Build a CharacterProperties object from a CHPX SPRM byte array.
    /// Simple grpprl parsing: for each SPRM opcode at index i, use byte[i+2] as operand.
    /// </summary>
    public static CharacterProperties FromChpx(byte[]? chpx)
    {
        var props = new CharacterProperties();
        if (chpx == null || chpx.Length < 3)
            return props;

        // Based on CHPX parsing in WordDocumentParser (parse opcode and operand) fileciteturn3file17
        for (int i = 0; i < chpx.Length - 2; i++)
        {
            byte sprm = chpx[i];
            byte val = chpx[i + 2];
            switch (sprm)
            {
                case 0x08: props.IsBold = (val & 1) != 0; break;
                case 0x09: props.IsItalic = (val & 1) != 0; break;
                case 0x0A: props.IsStrikeThrough = (val & 1) != 0; break;
                case 0x0D: props.IsSmallCaps = (val & 1) != 0; break;
                case 0x0E: props.IsAllCaps = (val & 1) != 0; break;
                case 0x0F: props.IsHidden = (val & 1) != 0; break;
                case 0x18: props.IsUnderlined = (val & 1) != 0; break;
                case 0x2A: props.FontSize = val; break;
                    // add other SPRM codes as required
            }
            // Skip over the operand byte as well
            i += 2;
        }
        return props;
    }
}
