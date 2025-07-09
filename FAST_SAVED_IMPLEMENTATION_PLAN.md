# Fast-Saved Document Implementation Plan for WvWareNet

## Overview

This document outlines a comprehensive plan to implement fast-saved document support in the WvWareNet C# parser. Fast-saved documents require complex piece table processing, cross-piece property boundary detection, and sophisticated text reconstruction algorithms.

## Current State Analysis

### What We Have
- Basic piece table parsing in `PieceTable.cs`
- Simple text extraction from pieces
- Rudimentary fast-saved document detection (lines 77-142 in `PieceTable.cs`)
- Basic CLX parsing

### What We're Missing
- Complex paragraph boundary detection across pieces
- Character property boundary detection
- FKP (Formatted Disk Page) parsing and integration
- Cross-piece property coordination
- Property modification (SPRM) processing from piece descriptors
- Complex document structure navigation

## Implementation Plan

### Phase 1: Core Infrastructure (Weeks 1-2)

#### 1.1 Create FKP (Formatted Disk Page) Infrastructure

**Files to Create:**
- `WvWareNet/Core/FormattedDiskPage.cs`
- `WvWareNet/Core/ParagraphFKP.cs` 
- `WvWareNet/Core/CharacterFKP.cs`
- `WvWareNet/Core/BTE.cs` (Bin Table Entry)

**Key Classes:**
```csharp
public class FormattedDiskPage
{
    public uint[] RgFc { get; set; }          // File coordinate array
    public byte[] RgBx { get; set; }          // Offset array to property structures
    public byte[] PropertyData { get; set; }  // Raw property data
}

public class ParagraphFKP : FormattedDiskPage
{
    public List<ParagraphProperties> GetParagraphProperties();
    public uint FindLargestFCLessThan(uint targetFC);
    public uint FindSmallestFCGreaterThan(uint targetFC);
}

public class CharacterFKP : FormattedDiskPage  
{
    public List<CharacterProperties> GetCharacterProperties();
    public uint FindLargestFCLessThan(uint targetFC);
    public uint FindSmallestFCGreaterThan(uint targetFC);
}

public class BTE
{
    public uint FC { get; set; }     // File coordinate
    public uint PageNumber { get; set; }  // FKP page number
}
```

#### 1.2 Implement Property Boundary Detection

**Files to Create:**
- `WvWareNet/Core/PropertyBoundaryDetector.cs`

**Key Methods:**
```csharp
public class PropertyBoundaryDetector
{
    public static (uint fcFirst, uint fcLim) GetComplexParaBounds(
        uint currentFC, CLX clx, BTE[] btePapx, uint[] posPapx, 
        int pieceIndex, Stream wordDocStream);
        
    public static (uint fcFirst, uint fcLim) GetComplexCharBounds(
        uint currentFC, CLX clx, BTE[] bteChpx, uint[] posChpx,
        int pieceIndex, Stream wordDocStream);
        
    private static uint GetComplexParaFcFirst(/* parameters */);
    private static uint GetComplexParaFcLim(/* parameters */);
    private static bool QuerySamePiece(uint fc, CLX clx, int pieceIndex);
}
```

#### 1.3 Enhance CLX and Piece Table Support

**Files to Modify:**
- `WvWareNet/Core/PieceTable.cs`
- `WvWareNet/Core/PieceDescriptor.cs`

**Enhancements:**
```csharp
public class CLX
{
    public PieceDescriptor[] Pcd { get; set; }
    public uint[] Pos { get; set; }              // Character position array
    public byte[][] Grpprl { get; set; }         // Property modification groups
    public uint[] CbGrpprl { get; set; }         // Sizes of Grpprl arrays
    public int Nopcd { get; set; }               // Number of pieces
}

public class PieceDescriptor
{
    // Existing properties...
    public PropertyModifier Prm { get; set; }    // Property modifications
    public bool HasComplexProperties => Prm?.FComplex ?? false;
}

public class PropertyModifier
{
    public bool FComplex { get; set; }
    public uint IgrpPrl { get; set; }    // Index into CLX.Grpprl for complex props
    public ushort Isprm { get; set; }    // Single property modifier for simple props
    public byte Val { get; set; }        // Property value for simple props
}
```

### Phase 2: Property Processing Infrastructure (Weeks 3-4)

#### 2.1 Implement SPRM (Style Property Modifier) Processing

**Files to Create:**
- `WvWareNet/Core/SPRM.cs`
- `WvWareNet/Core/SPRMProcessor.cs`

**Key Classes:**
```csharp
public class SPRM
{
    public ushort Operation { get; set; }
    public byte[] Operand { get; set; }
    public SPRMType Type { get; set; }
    public SPRMCategory Category { get; set; }
}

public enum SPRMCategory
{
    Paragraph = 1,
    Character = 2, 
    Picture = 3,
    Section = 4,
    Table = 5
}

public class SPRMProcessor
{
    public static void ApplyParagraphSPRMs(ParagraphProperties pap, 
        SPRM[] sprms, Stylesheet stsh);
    public static void ApplyCharacterSPRMs(CharacterProperties chp,
        SPRM[] sprms, Stylesheet stsh);
    public static SPRM[] ParseSPRMGroup(byte[] grpprl);
}
```

#### 2.2 Create Complex Property Assembly

**Files to Create:**
- `WvWareNet/Core/ComplexPropertyAssembler.cs`

**Key Methods:**
```csharp
public class ComplexPropertyAssembler
{
    public static ParagraphProperties AssembleComplexPAP(
        ParagraphProperties basePAP, int pieceIndex, CLX clx, 
        Stylesheet stsh);
        
    public static CharacterProperties AssembleComplexCHP(
        CharacterProperties baseCHP, int pieceIndex, CLX clx,
        Stylesheet stsh);
        
    private static SPRM[] GetPiecePropertyModifications(
        int pieceIndex, CLX clx);
}
```

#### 2.3 Implement Coordinate Conversion

**Files to Create:**
- `WvWareNet/Core/CoordinateConverter.cs`

**Key Methods:**
```csharp
public class CoordinateConverter
{
    public static uint ConvertCPToFC(uint characterPosition, CLX clx);
    public static uint ConvertFCToCP(uint fileCoordinate, CLX clx);
    public static int FindPieceForCP(uint characterPosition, CLX clx);
    public static int FindPieceForFC(uint fileCoordinate, CLX clx);
}
```

### Phase 3: Complex Document Parser (Weeks 5-7)

#### 3.1 Create Main Complex Document Parser

**Files to Create:**
- `WvWareNet/Parsers/ComplexDocumentParser.cs`

**Key Structure:**
```csharp
public class ComplexDocumentParser
{
    private readonly ILogger _logger;
    private readonly CLX _clx;
    private readonly BTE[] _btePapx;
    private readonly BTE[] _bteChpx;
    private readonly uint[] _posPapx;
    private readonly uint[] _posChpx;
    
    public DocumentModel ParseComplexDocument(FileInformationBlock fib, 
        byte[] wordDocStream, byte[] tableStream);
        
    private void ProcessPiece(int pieceIndex, Stream wordDocStream);
    private void HandlePropertyBoundaries(uint currentFC, uint currentCP);
    private void ProcessCharacter(char character, uint fc, uint cp);
}
```

#### 3.2 Implement State Management

**Files to Create:**
- `WvWareNet/Core/DocumentParseState.cs`

**Key Classes:**
```csharp
public class DocumentParseState
{
    // Current position tracking
    public uint CurrentCP { get; set; }
    public uint CurrentFC { get; set; }
    public int CurrentPiece { get; set; }
    
    // Property boundary tracking
    public uint ParaFcFirst { get; set; }
    public uint ParaFcLim { get; set; }
    public uint CharFcFirst { get; set; }
    public uint CharFcLim { get; set; }
    
    // Current properties
    public ParagraphProperties CurrentPAP { get; set; }
    public CharacterProperties CurrentCHP { get; set; }
    
    // Pending state
    public bool ParaPendingClose { get; set; }
    public bool CharPendingClose { get; set; }
    
    // Dirty flags
    public bool ParaDirty { get; set; }
    public bool CharDirty { get; set; }
}
```

#### 3.3 Enhance Document Model

**Files to Modify:**
- `WvWareNet/Core/DocumentModel.cs`
- `WvWareNet/Core/Paragraph.cs`
- `WvWareNet/Core/TextRun.cs`

**Enhancements:**
```csharp
public class Paragraph
{
    // Existing properties...
    public ParagraphProperties Properties { get; set; }
    public uint FcFirst { get; set; }
    public uint FcLim { get; set; }
    public List<TextRun> TextRuns { get; set; }
}

public class TextRun  
{
    public string Text { get; set; }
    public CharacterProperties Properties { get; set; }
    public uint FcFirst { get; set; }
    public uint FcLim { get; set; }
}
```

### Phase 4: Integration and Testing (Weeks 8-9)

#### 4.1 Integrate with Main Parser

**Files to Modify:**
- `WvWareNet/Parsers/WordDocumentParser.cs`
- `WvWareNet/WvDocExtractor.cs`

**Integration Points:**
```csharp
public class WordDocumentParser
{
    public DocumentModel ParseDocument(FileInformationBlock fib, 
        byte[] wordDocStream, byte[] tableStream)
    {
        if (fib.FComplex)
        {
            var complexParser = new ComplexDocumentParser(_logger);
            return complexParser.ParseComplexDocument(fib, wordDocStream, tableStream);
        }
        else
        {
            // Existing simple document parsing
            return ParseSimpleDocument(fib, wordDocStream);
        }
    }
}
```

#### 4.2 Create Comprehensive Test Suite

**Files to Create:**
- `WvWareNet.Tests/ComplexDocumentTests.cs`
- `WvWareNet.Tests/PropertyBoundaryTests.cs`
- `WvWareNet.Tests/SPRMProcessorTests.cs`
- `WvWareNet.Tests/FastSavedDocumentTests.cs`

**Test Categories:**
1. Basic complex document parsing
2. Property boundary detection across pieces
3. SPRM application and property assembly
4. Cross-piece text reconstruction
5. Fast-saved specific scenarios
6. Regression tests against existing functionality

### Phase 5: Advanced Features and Optimization (Weeks 10-12)

#### 5.1 Implement Advanced Property Features

**Features to Add:**
- Table detection and processing (`apap.fInTable`)
- Comment boundaries and processing
- Section properties across pieces
- List numbering continuation across pieces

#### 5.2 Performance Optimization

**Optimization Areas:**
- FKP caching to avoid repeated parsing
- Property boundary caching
- Efficient piece lookup algorithms
- Memory usage optimization for large documents

#### 5.3 Error Handling and Robustness

**Robustness Features:**
- Graceful degradation when property parsing fails
- Fallback to simple parsing for malformed complex documents
- Comprehensive logging for debugging complex documents
- Stream bounds checking and validation

## Technical Implementation Details

### Property Boundary Detection Algorithm

The core algorithm follows wvWare's `wvGetComplexParaBounds`:

```csharp
public static (uint fcFirst, uint fcLim) GetComplexParaBounds(
    uint currentFC, CLX clx, BTE[] btePapx, uint[] posPapx, 
    int pieceIndex, Stream wordDocStream)
{
    // 1. Find BTE entry for current FC
    var bte = FindBTEFromFC(currentFC, btePapx, posPapx);
    
    // 2. Load FKP page for this BTE
    var fkp = LoadParagraphFKP(bte.PageNumber, wordDocStream);
    
    // 3. Find paragraph boundaries
    uint fcFirst = GetComplexParaFcFirst(currentFC, clx, btePapx, 
        posPapx, pieceIndex, fkp, wordDocStream);
    uint fcLim = GetComplexParaFcLim(currentFC, clx, btePapx,
        posPapx, pieceIndex, fkp, wordDocStream);
        
    return (fcFirst, fcLim);
}
```

### Cross-Piece Property Search

Following wvWare's approach for searching across pieces:

```csharp
private static uint GetComplexParaFcFirst(/* parameters */)
{
    uint fcTest = fkp.FindLargestFCLessThan(currentFC);
    
    if (QuerySamePiece(fcTest - 1, clx, pieceIndex))
    {
        return fcTest - 1;
    }
    else
    {
        // Search backwards through pieces
        int searchPiece = pieceIndex - 1;
        while (searchPiece >= 0)
        {
            uint endFC = GetEndFCPiece(searchPiece, clx);
            var searchBTE = FindBTEFromFC(endFC, btePapx, posPapx);
            var searchFKP = LoadParagraphFKP(searchBTE.PageNumber, wordDocStream);
            
            fcTest = searchFKP.FindLargestFCLessThan(endFC);
            if (QuerySamePiece(fcTest - 1, clx, searchPiece))
            {
                return fcTest - 1;
            }
            searchPiece--;
        }
        
        // Fallback to current FC if no paragraph boundary found
        return currentFC;
    }
}
```

## Integration with Existing Code

### Minimal Changes to Existing Classes

1. **PieceTable.cs**: Add CLX parsing and enhance piece descriptor
2. **WordDocumentParser.cs**: Add complex document detection and delegation
3. **WvDocExtractor.cs**: Route complex documents to new parser
4. **DocumentModel.cs**: Enhance with property information

### Backward Compatibility

- Existing simple document parsing remains unchanged
- New complex parsing is additive
- All existing tests should continue to pass
- Performance impact only affects complex documents

## Success Criteria

### Functional Requirements
- [ ] Parse fast-saved documents with multiple pieces correctly
- [ ] Extract text in proper order across piece boundaries
- [ ] Detect paragraph boundaries accurately
- [ ] Apply character and paragraph properties correctly
- [ ] Handle property modifications from piece descriptors
- [ ] Process all example fast-saved documents successfully

### Performance Requirements
- [ ] Parse complex documents within 2x time of simple parsing
- [ ] Memory usage scales linearly with document size
- [ ] No regression in simple document parsing performance

### Quality Requirements
- [ ] Comprehensive test coverage (>90%)
- [ ] All existing tests continue to pass
- [ ] Robust error handling and logging
- [ ] Clear documentation and examples

## Risk Mitigation

### Technical Risks
1. **Complexity**: Implement incrementally with extensive testing
2. **Performance**: Profile and optimize critical paths
3. **Compatibility**: Maintain backward compatibility through feature flags
4. **Quality**: Extensive testing with diverse document samples

### Schedule Risks
- Buffer time built into each phase
- Parallel development where possible
- Early integration testing to catch issues

## Conclusion

This implementation plan provides a comprehensive roadmap for adding fast-saved document support to WvWareNet. The phased approach ensures manageable development while maintaining quality and compatibility. The resulting implementation will handle the complex scenarios that the current parser cannot, bringing it to parity with the original wvWare functionality.