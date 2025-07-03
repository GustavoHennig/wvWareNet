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

    /// <summary>
    /// The starting character position (cp) for this piece
    /// </summary>
    public int CpStart { get; set; }

    /// <summary>
    /// The ending character position (cp) for this piece
    /// </summary>
    public int CpEnd { get; set; }

    /// <summary>
    /// Reference to the associated CHPX (character formatting) data
    /// </summary>
    public byte[] Chpx { get; set; }

    /// <summary>
    /// The starting file position (FC) for this piece
    /// </summary>
    public int FcStart { get; set; }

    /// <summary>
    /// The ending file position (FC) for this piece
    /// </summary>
    public int FcEnd { get; set; }

    public static PieceDescriptor Parse(byte[] bytes)
    {
        var descriptor = new PieceDescriptor();
        using (var stream = new MemoryStream(bytes))
        using (var reader = new BinaryReader(stream))
        {
            // First 2 bytes are reserved
            reader.ReadBytes(2);
            uint fc = reader.ReadUInt32();
            descriptor.IsUnicode = (fc & 0x40000000) != 0;
            descriptor.FilePosition = (fc & 0x3FFFFFFF); // Mask out the Unicode flag
            // Next 2 bytes are PRM, which we can ignore for now
            reader.ReadBytes(2);
        }
        return descriptor;
    }
}
