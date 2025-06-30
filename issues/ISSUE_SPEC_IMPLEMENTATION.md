# Implementation of Missing MS-DOC Spec Requirements

The current parser needs enhancements to fully implement the MS-DOC specification for text retrieval and paragraph boundaries. While core functionality exists, some edge cases and specific requirements from the spec are not completely handled.

## Missing Spec Requirements

1. **Special FcCompressed.fc Values Handling**
   - The spec mentions special values in the table defined in FcCompressed.fc description (step 6 of text retrieval algorithm)
   - Current implementation doesn't explicitly handle these special cases

2. **Strict Character Position Validation**
   - Negative character positions should be explicitly rejected
   - Character position ranges should be strictly validated against document bounds

3. **Paragraph Boundary Edge Cases**
   - Some edge cases in paragraph boundary detection could be more thoroughly implemented
   - Better handling of TTP marks and cell marks as paragraph boundaries

4. **Complete PlcPcd Implementation**
   - Current PieceTable handles basic cases but could better match the spec's requirements
   - More robust handling of invalid/malformed piece tables

## Reference Implementations

1. **wvWare C Implementation**
   - `text.c`: Contains core text retrieval logic
   - `decode_complex.c`: Handles complex document structures
   - `clx.c`: Implements CLX parsing

2. **OnlyOffice C++ Implementation**
   - `MsBinaryFileReader.cpp`: Alternative implementation of DOC parsing
   - Handles some edge cases differently than wvWare

## Execution Plan

1. **Analyze Spec Requirements**
   - Carefully review sections 2.4.1 and 2.4.2 of MS-DOC spec
   - Identify all edge cases and special scenarios

2. **Enhance FcCompressed Handling**
   - Implement special value handling from spec
   - Add validation for compressed/uncompressed flags
   - Unit tests for all special cases

3. **Improve Position Validation**
   - Add explicit checks for negative positions
   - Validate character positions against document bounds
   - Add boundary condition tests

4. **Enhance Paragraph Detection**
   - Implement complete paragraph mark handling
   - Better detection of TTP marks and cell marks
   - Add tests for complex paragraph scenarios

5. **Robust PieceTable Handling**
   - Add more validation for piece table data
   - Better error recovery for malformed tables
   - Additional logging for debugging

6. **Testing Strategy**
   - Create test documents for all edge cases
   - Add unit tests for each spec requirement
   - Verify against reference implementations

## Implementation Details

Key files to modify:
- `WvWareNet/Core/PieceTable.cs`: Enhance piece table handling
- `WvWareNet/Parsers/WordDocumentParser.cs`: Improve text retrieval
- `WvWareNet/Core/FileInformationBlock.cs`: Add validation

## Important Notes
- Must maintain backward compatibility
- No hardcoded solutions - handle all cases generically
- Avoid reflection
- Ensure performance isn't degraded


## Official MS-DOC Specification Reference

2.4.1 Retrieving Text
The following algorithm specifies how to find the text at a particular character position (cp). Negative
character positions are not valid.
1. Read the FIB from offset zero in the WordDocument Stream.

2. All versions of the FIB contain exactly one FibRgFcLcb97, though it can be nested in a larger
structure. FibRgFcLcb97.fcClx specifies the offset in the Table Stream of a Clx.
FibRgFcLcb97.lcbClx specifies the size, in bytes, of that Clx. Read the Clx from the Table
Stream.
3. The Clx contains a Pcdt, and the Pcdt contains a PlcPcd. Find the largest i such that
PlcPcd.aCp[i] ≤ cp. As with all Plcs, the elements of PlcPcd.aCp are sorted in ascending order.
Recall from the definition of a Plc that the aCp array has one more element than the aPcd array.
Thus, if the last element of PlcPcd.aCp is less than or equal to cp, cp is outside the range of valid
character positions in this document.
4. PlcPcd.aPcd[i] is a Pcd. Pcd.fc is an FcCompressed that specifies the location in the
WordDocument Stream of the text at character position PlcPcd.aCp[i].
5. If FcCompressed.fCompressed is zero, the character at position cp is a 16-bit Unicode
character at offset FcCompressed.fc + 2(cp - PlcPcd.aCp[i]) in the WordDocument Stream.
This is to say that the text at character position PlcPcd.aCP[i] begins at offset
FcCompressed.fc in the WordDocument Stream and each character occupies two bytes.
6. If FcCompressed.fCompressed is 1, the character at position cp is an 8-bit ANSI character at
offset (FcCompressed.fc / 2) + (cp - PlcPcd.aCp[i]) in the WordDocument Stream, unless it is
one of the special values in the table defined in the description of FcCompressed.fc. This is to
say that the text at character position PlcPcd.aCP[i] begins at offset FcCompressed.fc / 2 in
the WordDocument Stream and each character occupies one byte.
2.4.2 Determining Paragraph Boundaries
This section specifies how to find the beginning and end character positions of the paragraph that
contains a given character position. The character at the end character position of a paragraph MUST
be a paragraph mark, an end-of-section character, a cell mark, or a TTP mark (See Overview of
Tables). Negative character positions are not valid.
To find the character position of the first character in the paragraph that contains a given character
position cp:
1. Follow the algorithm from Retrieving Text up to and including step 3 to find i. Also remember the
FibRgFcLcb97 and PlcPcd found in step 1 of Retrieving Text. If the algorithm from Retrieving
Text specifies that cp is invalid, leave the algorithm.
2. Let pcd be PlcPcd.aPcd[i].
3. Let fcPcd be Pcd.fc.fc. Let fc be fcPcd + 2(cp – PlcPcd.aCp[i]). If Pcd.fc.fCompressed is one,
set fc to fc / 2, and set fcPcd to fcPcd/2.
4. Read a PlcBtePapx at offset FibRgFcLcb97.fcPlcfBtePapx in the Table Stream, and of size
FibRgFcLcb97.lcbPlcfBtePapx. Let fcLast be the last element of plcbtePapx.aFc. If fcLast is
less than or equal to fc, examine fcPcd. If fcLast is less than fcPcd, go to step 8. Otherwise, set
fc to fcLast. If Pcd.fc.fCompressed is one, set fcLast to fcLast / 2. Set fcFirst to fcLast and
go to step 7.
5. Find the largest j such that plcbtePapx.aFc[j] ≤ fc. Read a PapxFkp at offset
aPnBtePapx[j].pn *512 in the WordDocument Stream.
6. Find the largest k such that PapxFkp.rgfc[k] ≤ fc. If the last element of PapxFkp.rgfc is less
than or equal to fc, then cp is outside the range of character positions in this document, and is
not valid. Let fcFirst be PapxFkp.rgfc[k].
7. If fcFirst is greater than fcPcd, then let dfc be (fcFirst – fcPcd). If Pcd.fc.fCompressed is
zero, then set dfc to dfc / 2. The first character of the paragraph is at character position
PlcPcd.aCp[i] + dfc. Leave the algorithm.
8. If PlcPcd.aCp[i] is 0, then the first character of the paragraph is at character position 0. Leave
the algorithm.
9. Set cp to PlcPcd.aCp[i]. Set i to i - 1. Go to step 2.
To find the character position of the last character in the paragraph that contains a given character
position cp:
1. Follow the algorithm from Retrieving Text up to and including step 3 to find i. Also remember the
FibRgFcLcb97, and PlcPcd found in step 1 of Retrieving Text. If the algorithm from Retrieving
Text specifies that cp is invalid, leave the algorithm.
2. Let pcd be PlcPcd.aPcd[i].
3. Let fcPcd be Pcd.fc.fc. Let fc be fcPcd + 2(cp – PlcPcd.aCp[i]). Let fcMac be fcPcd +
2(PlcPcd.aCp[i+1] - PlcPcd.aCp[i]). If Pcd.fc.fCompressed is one, set fc to fc/2, set fcPcd to
fcPcd /2 and set fcMac to fcMac/2.
4. Read a PlcBtePapx at offset FibRgFcLcb97.fcPlcfBtePapx in the Table Stream, and of size
FibRgFcLcb97.lcbPlcfBtePapx. Then find the largest j such that plcbtePapx.aFc[j] ≤ fc. If the
last element of plcbtePapx.aFc is less than or equal to fc, then go to step 7. Read a PapxFkp at
offset aPnBtePapx[j].pn *512 in the WordDocument Stream.
5. Find largest k such that PapxFkp.rgfc[k] ≤ fc. If the last element of PapxFkp.rgfc is less than
or equal to fc, then cp is outside the range of character positions in this document, and is not
valid. Let fcLim be PapxFkp.rgfc[k+1].
6. If fcLim ≤ fcMac, then let dfc be (fcLim – fcPcd). If Pcd.fc.fCompressed is zero, then set dfc
to dfc / 2. The last character of the paragraph is at character position PlcPcd.aCp[i] + dfc – 1.
Leave the algorithm.
7. Set cp to PlcPcd.aCp[i+1]. Set i to i + 1. Go to step 2.