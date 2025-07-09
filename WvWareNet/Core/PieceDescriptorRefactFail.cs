using System.IO;

namespace WvWareNet.Core
{
    public class PieceDescriptorRefactFail
    {
        /// <summary>
        /// The raw FC value as read from the file (before normalization)
        /// </summary>
        public uint RawFc { get; set; }

        /// <summary>
        /// The PRM (Property Modifier) value
        /// </summary>
        public uint Prm { get; set; }

        /// <summary>
        /// The file position (FC) where this piece starts (after normalization)
        /// </summary>
        public int FcStart { get; set; }

        /// <summary>
        /// Specifies whether the text in this piece is stored in Unicode (UTF-16LE)
        /// </summary>
        public bool IsUnicode { get; set; }

        /// <summary>
        /// The starting character position (CP) for this piece
        /// </summary>
        public int CpStart { get; set; }

        /// <summary>
        /// The ending character position (CP) for this piece
        /// </summary>
        public int CpEnd { get; set; }

        /// <summary>
        /// Reference to the associated CHPX (character formatting) data
        /// </summary>
        public byte[]? Chpx { get; set; }

        /// <summary>
        /// The ending file position (FC) for this piece (calculated)
        /// </summary>
        public int FcEnd => FcStart + (IsUnicode ? (CpEnd - CpStart) * 2 : (CpEnd - CpStart));

        /// <summary>
        /// Indicates if the piece contains formatting information
        /// </summary>
        public bool HasFormatting { get; set; }

        /// <summary>
        /// Reserved bits from the original format
        /// </summary>
        public byte ReservedFlags { get; set; }

        /// <summary>
        /// The file position (FC) where this piece starts (legacy property for compatibility)
        /// </summary>
        public uint FilePosition
        {
            get => (uint)FcStart;
            set => FcStart = (int)value;
        }

        public static PieceDescriptor Parse(byte[] bytes)
        {
            var descriptor = new PieceDescriptor();
            using (var stream = new MemoryStream(bytes))
            using (var reader = new BinaryReader(stream))
            {
                // First 2 bytes are reserved
                reader.ReadBytes(2);
                uint fc = reader.ReadUInt32();
                descriptor.RawFc = fc;
                descriptor.IsUnicode = (fc & 0x40000000) == 0;
                descriptor.FcStart = (int)(fc & 0x3FFFFFFF);
                // Next 2 bytes are PRM
                descriptor.Prm = reader.ReadUInt16();
            }
            return descriptor;
        }
    }
}