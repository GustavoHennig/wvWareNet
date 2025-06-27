namespace WvWareNet.Core;

public class FileInformationBlock
{
    // Document identification and version
    public ushort WIdent { get; set; }
    public ushort NFib { get; set; }
    public ushort NProduct { get; set; }
    public string? FibVersion
    {
        get
        {
            return NFib switch
            {
                0x0062 => "Word 6",
                0x0063 => "Word 6.1",
                0x0065 => "Word 95",
                0x0076 => "Word 97",
                0x00C1 => "Word 97/2000",
                0x00D9 => "Word 2000",
                0x0101 => "Word 2003",
                0x0112 => "Word 2007",
                _ => null
            };
        }
    }
    
    // Language ID
    public ushort Lid { get; set; }
    
    // Document properties
    public short PnNext { get; set; }
    public bool FDot { get; set; }
    public bool FGlsy { get; set; }
    public bool FComplex { get; set; }
    public bool FHasPic { get; set; }
    public byte CQuickSaves { get; set; }
    public bool FEncrypted { get; set; }
    public bool FWhichTblStm { get; set; }
    public bool FReadOnlyRecommended { get; set; }
    public bool FWriteReservation { get; set; }
    public bool FExtChar { get; set; }
    public bool FLoadOverride { get; set; }
    public bool FFarEast { get; set; }
    public bool FCrypto { get; set; }
    public ushort NFibBack { get; set; }
    public uint LKey { get; set; }
    
    // Environment information
    public byte Envr { get; set; }
    public bool FMac { get; set; }
    public bool FEmptySpecial { get; set; }
    public bool FLoadOverridePage { get; set; }
    public bool FFutureSavedUndo { get; set; }
    public bool FWord97Saved { get; set; }
    public byte FSpare0 { get; set; }
    
    // Character set handling
    public ushort Chse { get; set; }
    public ushort ChsTables { get; set; }
    
    // File position markers
    public uint FcMin { get; set; }
    public uint FcMac { get; set; }
    public ushort Csw { get; set; }
    
    // Magic numbers
    public ushort WMagicCreated { get; set; }
    public ushort WMagicRevised { get; set; }
    public ushort WMagicCreatedPrivate { get; set; }
    public ushort WMagicRevisedPrivate { get; set; }
    
    // Word 6 specific properties
    public short PnFbpChpFirst_W6 { get; set; }
    public short PnChpFirst_W6 { get; set; }
    public short CpnBteChp_W6 { get; set; }
    public short PnFbpPapFirst_W6 { get; set; }
    public short PnPapFirst_W6 { get; set; }
    public short CpnBtePap_W6 { get; set; }
    public short PnFbpLvcFirst_W6 { get; set; }
    public short PnLvcFirst_W6 { get; set; }
    public short CpnBteLvc_W6 { get; set; }
    
    // Far East language ID
    public short LidFE { get; set; }
    
    // Document statistics
    public ushort Clw { get; set; }
    public int CbMac { get; set; }
    public uint LProductCreated { get; set; }
    public uint LProductRevised { get; set; }
    
    // Character count properties
    public uint CcpText { get; set; }
    public int CcpFtn { get; set; }
    public int CcpHdr { get; set; }
    public int CcpMcr { get; set; }
    public int CcpAtn { get; set; }
    public int CcpEdn { get; set; }
    public int CcpTxbx { get; set; }
    public int CcpHdrTxbx { get; set; }
    
    // Formatting properties
    public int PnFbpChpFirst { get; set; }
    public int PnChpFirst { get; set; }
    public int CpnBteChp { get; set; }
    public int PnFbpPapFirst { get; set; }
    public int PnPapFirst { get; set; }
    public int CpnBtePap { get; set; }
    public int PnFbpLvcFirst { get; set; }
    public int PnLvcFirst { get; set; }
    public int CpnBteLvc { get; set; }
    
    // Island properties
    public int FcIslandFirst { get; set; }
    public int FcIslandLim { get; set; }
    
    // Count of FCLCB structures
    public ushort Cfclcb { get; set; }
    
    // File position and length properties
    // (Only including key properties for text extraction)
    public int FcStshf { get; set; }
    public uint LcbStshf { get; set; }
    public int FcPlcfbteChpx { get; set; }
    public uint LcbPlcfbteChpx { get; set; }
    public int FcPlcfbtePapx { get; set; }
    public uint LcbPlcfbtePapx { get; set; }
    public int FcClx { get; set; }
    public uint LcbClx { get; set; }
    
    // Additional properties would go here
    // (Omitted for brevity in initial implementation)
    
    public FileInformationBlock()
    {
        // Initialize all properties to default values
        WIdent = 0;
        NFib = 0;
        NProduct = 0;
        Lid = 0;
        PnNext = 0;
        FDot = false;
        FGlsy = false;
        FComplex = false;
        FHasPic = false;
        CQuickSaves = 0;
        FEncrypted = false;
        FWhichTblStm = false;
        FReadOnlyRecommended = false;
        FWriteReservation = false;
        FExtChar = false;
        FLoadOverride = false;
        FFarEast = false;
        FCrypto = false;
        NFibBack = 0;
        LKey = 0;
        Envr = 0;
        FMac = false;
        FEmptySpecial = false;
        FLoadOverridePage = false;
        FFutureSavedUndo = false;
        FWord97Saved = false;
        FSpare0 = 0;
        Chse = 0;
        ChsTables = 0;
        FcMin = 0;
        FcMac = 0;
        Csw = 0;
        WMagicCreated = 0;
        WMagicRevised = 0;
        WMagicCreatedPrivate = 0;
        WMagicRevisedPrivate = 0;
        // ... (all other properties initialized to 0/false)
    }

    // Offsets for header/footer/footnote PLCFs (Word 97+)
    public int FcPlcfhdd { get; set; }
    public uint LcbPlcfhdd { get; set; }
    public int FcPlcfftn { get; set; }
    public uint LcbPlcfftn { get; set; }

    // Text Box properties
    public int FcPlcftxbxTxt { get; set; }
    public uint LcbPlcftxbxTxt { get; set; }

        public static FileInformationBlock Parse(byte[] wordDocumentStream)
        {
            var fib = new FileInformationBlock();
            if (wordDocumentStream.Length < 512) // A minimal FIB requires at least the header size
            {
                return fib; // Return a default FIB if stream is too small
            }

            using var ms = new System.IO.MemoryStream(wordDocumentStream);
            using var reader = new System.IO.BinaryReader(ms);

            // Read basic FIB fields from the beginning
            fib.WIdent = reader.ReadUInt16();
            fib.NFib = reader.ReadUInt16();
            fib.NProduct = reader.ReadUInt16();
            fib.Lid = reader.ReadUInt16();
            fib.PnNext = reader.ReadInt16();

            // Parse flag bits
            ushort flags = reader.ReadUInt16();
            fib.FDot = (flags & 0x0001) != 0;
            fib.FGlsy = (flags & 0x0002) != 0;
            fib.FComplex = (flags & 0x0004) != 0;
            fib.FHasPic = (flags & 0x0008) != 0;
            fib.CQuickSaves = (byte)((flags >> 4) & 0x000F);
            fib.FEncrypted = (flags & 0x0100) != 0;
            fib.FWhichTblStm = (flags & 0x0200) != 0;
            fib.FReadOnlyRecommended = (flags & 0x0400) != 0;
            fib.FWriteReservation = (flags & 0x0800) != 0;
            fib.FExtChar = (flags & 0x1000) != 0;
            fib.FLoadOverride = (flags & 0x2000) != 0;
            fib.FFarEast = (flags & 0x4000) != 0;
            fib.FCrypto = (flags & 0x8000) != 0;

            fib.NFibBack = reader.ReadUInt16();
            fib.LKey = reader.ReadUInt32();

            fib.Envr = reader.ReadByte();
            byte envrFlags = reader.ReadByte();
            fib.FMac = (envrFlags & 0x01) != 0;
            fib.FEmptySpecial = (envrFlags & 0x02) != 0;
            fib.FLoadOverridePage = (envrFlags & 0x04) != 0;
            fib.FFutureSavedUndo = (envrFlags & 0x08) != 0;
            fib.FWord97Saved = (envrFlags & 0x10) != 0;
            fib.FSpare0 = (byte)(envrFlags >> 5);

            fib.Chse = reader.ReadUInt16();
            fib.ChsTables = reader.ReadUInt16();

            // Read FcMin and FcMac at standard offset
            ms.Position = 0x18; 
            fib.FcMin = reader.ReadUInt32();
            fib.FcMac = reader.ReadUInt32();

            // Determine version-specific offsets
            int offsetPlcfbteChpx = 0x00FA; // Default for Word 6/95
            int offsetPlcfbtePapx = 0x0102;
            int offsetClx = 0x00A4;
            int offsetPlcfhdd = 0x00F2;
            int offsetPlcfftn = 0x012A;
            int offsetPlcftxbxTxt = 0x01F6;

            if (fib.NFib >= 104) // Word 97 and later
            {
                offsetPlcfbteChpx = 0x014E;
                offsetPlcfbtePapx = 0x0156;
                offsetClx = 0x01A2;
                offsetPlcfhdd = 0x0142;
                offsetPlcfftn = 0x015E;
                offsetPlcftxbxTxt = 0x01C6;
            }

            // Read version-specific properties
            ms.Position = offsetPlcftxbxTxt;
            fib.FcPlcftxbxTxt = reader.ReadInt32();
            ms.Position = offsetPlcftxbxTxt + 4;
            fib.LcbPlcftxbxTxt = reader.ReadUInt32();

            ms.Position = offsetPlcfbteChpx;
            fib.FcPlcfbteChpx = reader.ReadInt32();
            fib.LcbPlcfbteChpx = reader.ReadUInt32();

            ms.Position = offsetPlcfbtePapx;
            fib.FcPlcfbtePapx = reader.ReadInt32();
            fib.LcbPlcfbtePapx = reader.ReadUInt32();

            ms.Position = offsetClx;
            fib.FcClx = reader.ReadInt32();
            fib.LcbClx = reader.ReadUInt32();

            ms.Position = offsetPlcfhdd;
            fib.FcPlcfhdd = reader.ReadInt32();
            fib.LcbPlcfhdd = reader.ReadUInt32();

            ms.Position = offsetPlcfftn;
            fib.FcPlcfftn = reader.ReadInt32();
            fib.LcbPlcfftn = reader.ReadUInt32();

            return fib;
        }
}
