namespace WvWareNet.Core;

public class PieceDescriptor
{
    /// <summary>
    /// The file position (FC) where this piece starts
    /// </summary>
    public uint FilePosition { get; set; }

    /// <summary>
    /// Specifies whether the text in this piece is stored in ANSI or Unicode
    /// </summary>
    public bool IsUnicode { get; set; }

    /// <summary>
    /// Indicates if the piece contains formatting information
    /// </summary>
    public bool HasFormatting { get; set; }

    /// <summary>
    /// Reserved bits from the original format
    /// </summary>
    public byte ReservedFlags { get; set; }
}
