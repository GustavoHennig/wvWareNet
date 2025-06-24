namespace WvWareNet.Core;

public class FileInformationBlock
{
    // Document identification and version
    public ushort WIdent { get; set; }
    public ushort NFib { get; set; }
    public ushort NProduct { get; set; }
    
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
}
