You have two implementations of the same functionality:

- A C version that correctly extracts pieces and computes fcStart/fcEnd by converting character positions (CP) to file offsets (FC).
- A C# version that uses a new PieceTable but currently computes fcStart and fcEnd incorrectly (for edge scenarios).

Task:

- Carefully compare the C and C# implementations of piece extraction and CP→FC conversion.
- Identify and fix any errors in WordDocumentParser.cs so the C# parser produces exactly the same output as the C version.
- Adapt WordDocumentParser.cs to your updated PieceTable implementation, making the minimum code changes necessary to compile and run correctly on the first try.
- Pay special attention to preventing infinite loops.
- Keep in mind the C version is able to extract text from all documents (exactly my goal), the C# do not.
- Be prapared for scenarios were fib values are insufficient or missing, such as FcPlcfbteChpx, LcbPlcfbteChpx, FcStshf, LcbStshf, etc. Make a logic to guess all values that are not extrictly mandatory for text extraction.

You must be prepared for insuficient data, you need to guess all data guessable,

After coding, do a final check of the new versions to confirm they work correctly for text extraction of binary Word documents.

Deliverable:
- A complete, revised WordDocumentParser.cs source file that meets these requirements.
- A complete, revised PieceTable.cs source file that meets these requirements.
- A brief summary of the changes you made and why.

```csharp
// NEW GENERATED CODE WITH A FIX PROPOSAL
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WvWareNet.Utilities;

namespace WvWareNet.Core
{
    public class PieceTable
    {
        private readonly ILogger _logger;
        private readonly List<PieceDescriptor> _pieces = new();
        private int[] _cpArray = Array.Empty<int>();

        public IReadOnlyList<PieceDescriptor> Pieces => _pieces;

        public PieceTable(ILogger logger)
        {
            _logger = logger;
        }

        public static PieceTable CreateFromStreams(ILogger logger, FileInformationBlock fib, byte[]? tableStream, byte[] wordDocStream)
        {
            var pieceTable = new PieceTable(logger);
            byte[]? clxData = null;

            if (tableStream != null && fib.LcbClx > 0 && fib.FcClx >= 0 && (fib.FcClx + fib.LcbClx) <= tableStream.Length)
            {
                clxData = new byte[fib.LcbClx];
                Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);
                pieceTable.Parse(clxData, fib);
            }
            
            if (pieceTable.Pieces.Count == 0)
            {
                logger.LogWarning("CLX data not found or failed to parse. Treating document as a single piece.");
                pieceTable.SetSinglePiece(fib);
            }

            return pieceTable;
        }

        public void Parse(byte[] clxData, FileInformationBlock fib)
        {
            _pieces.Clear();
            if (clxData.Length == 0) return;

            int plcPcdOffset = -1;
            int plcPcdLength = -1;

            int i = 0;
            while (i < clxData.Length)
            {
                byte clxType = clxData[i];
                if (clxType == 0x02) 
                {
                    if (i + 5 > clxData.Length) break;
                    plcPcdLength = BitConverter.ToInt32(clxData, i + 1);
                    plcPcdOffset = i + 5;
                    break;
                }
                if (clxType == 0x01)
                {
                    if (i + 3 > clxData.Length) break;
                    ushort prcSize = BitConverter.ToUInt16(clxData, i + 1);
                    i += 1 + 2 + prcSize;
                }
                else
                {
                    i++;
                }
            }

            if (plcPcdOffset == -1) return;

            try
            {
                using var stream = new MemoryStream(clxData, plcPcdOffset, plcPcdLength);
                using var reader = new BinaryReader(stream);

                int pieceCount = (plcPcdLength - 4) / 12;
                if (pieceCount <= 0) return;

                _cpArray = new int[pieceCount + 1];
                for (int j = 0; j <= pieceCount; j++)
                {
                    _cpArray[j] = reader.ReadInt32();
                }

                for (int j = 0; j < pieceCount; j++)
                {
                    uint fc = reader.ReadUInt32();
                    uint prm = reader.ReadUInt32();

                    _pieces.Add(new PieceDescriptor { RawFc = fc, Prm = prm });
                }

                if (!fib.FExtChar)
                {
                    _logger.LogInfo("[DEBUG] Applying Word 6/95 FC normalization logic.");
                    for (int j = 0; j < _pieces.Count; j++)
                    {
                        _pieces[j].RawFc = (_pieces[j].RawFc * 2) | 0x40000000U;
                    }
                }
                
                for (int j = 0; j < _pieces.Count; j++)
                {
                    uint rawFc = _pieces[j].RawFc;
                    _pieces[j].IsUnicode = (rawFc & 0x40000000) == 0;
                    _pieces[j].FcStart = (int)((rawFc & 0x3FFFFFFF) / (_pieces[j].IsUnicode ? 1 : 2));
                    _pieces[j].CpStart = _cpArray[j];
                    _pieces[j].CpEnd = _cpArray[j+1];
                }

                _logger.LogInfo($"Parsed {_pieces.Count} piece descriptors from piece table.");
            }
            catch (Exception ex)
            {
                _logger.LogError("Error parsing piece table data", ex);
                _pieces.Clear();
            }
        }
        
        public int GetPieceIndexFromCp(int cp)
        {
            for (int i = 0; i < _pieces.Count; i++)
            {
                if (cp >= _cpArray[i] && cp < _cpArray[i + 1])
                {
                    return i;
                }
            }
            if (_cpArray.Length > 0 && cp == _cpArray.Last())
            {
                return _pieces.Count - 1;
            }
            return -1;
        }
        
        public int ConvertCpToFc(int cp)
        {
            int pieceIndex = GetPieceIndexFromCp(cp);
            if (pieceIndex == -1)
            {
                _logger.LogWarning($"Could not find piece for CP={cp}");
                return -1;
            }

            var piece = _pieces[pieceIndex];
            int offsetInPiece = cp - piece.CpStart;
            
            return piece.FcStart + (piece.IsUnicode ? offsetInPiece * 2 : offsetInPiece);
        }

        public void SetSinglePiece(FileInformationBlock fib)
        {
            _pieces.Clear();
            _cpArray = new int[] { 0, (int)fib.CcpText };
            
            bool isUnicode = false; 
            int fcStart = (int)fib.FcMin;

            _pieces.Add(new PieceDescriptor
            {
                IsUnicode = isUnicode,
                CpStart = 0,
                CpEnd = _cpArray[1],
                FcStart = fcStart,
                RawFc = (uint)fcStart | 0x40000000U // Assume 8-bit for fallback
            });
        }
        
        public string GetTextForRange(int startCp, int endCp, Stream documentStream)
        {
            if (endCp <= startCp) return string.Empty;

            var sb = new System.Text.StringBuilder();
            int currentCp = startCp;

            while (currentCp < endCp)
            {
                int pieceIndex = GetPieceIndexFromCp(currentCp);
                if (pieceIndex == -1)
                {
                    _logger.LogWarning($"Could not find piece for CP={currentCp}. Stopping text extraction for range.");
                    break;
                }
                
                var piece = _pieces[pieceIndex];
                int cpInPieceStart = currentCp - piece.CpStart;
                int cpInPieceEnd = Math.Min(endCp, piece.CpEnd) - piece.CpStart;
                int count = cpInPieceEnd - cpInPieceStart;

                int byteOffset = piece.IsUnicode ? cpInPieceStart * 2 : cpInPieceStart;
                int byteCount = piece.IsUnicode ? count * 2 : count;

                if (piece.FcStart + byteOffset + byteCount > documentStream.Length)
                {
                    _logger.LogWarning("Text range extends beyond document stream. Truncating.");
                    byteCount = (int)documentStream.Length - (piece.FcStart + byteOffset);
                    if (byteCount < 0) byteCount = 0;
                }

                if (byteCount > 0)
                {
                    documentStream.Seek(piece.FcStart + byteOffset, SeekOrigin.Begin);
                    var reader = new BinaryReader(documentStream);
                    byte[] bytes = reader.ReadBytes(byteCount);

                    sb.Append(piece.IsUnicode
                        ? System.Text.Encoding.Unicode.GetString(bytes)
                        : System.Text.Encoding.GetEncoding(1252).GetString(bytes));
                }
                
                currentCp += count;
            }

            return sb.ToString();
        }
    }

}
```


```csharp
// WordDocumentParser.cs (That works with the most documents, currently the head version)
using System;
using System.Collections.Generic;
using System.Linq;
using WvWareNet.Parsers;
using WvWareNet.Utilities;

namespace WvWareNet.Parsers
{
    public class WordDocumentParser
    {
        private readonly CompoundFileBinaryFormatParser _cfbfParser;
        private readonly ILogger _logger;

        // Implements [MS-DOC] 2.4.2 paragraph boundary algorithm
        // Returns the character position of the first character in the paragraph containing cp
        private int FindParagraphStartCp(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream)
        {
            // Implements steps 1-3 of [MS-DOC] 2.4.2 for finding paragraph start
            // 1. Find i such that PlcPcd.aCp[i] <= cp < PlcPcd.aCp[i+1]
            // 2. Let pcd = PlcPcd.aPcd[i]
            // 3. Let fcPcd = pcd.fc.fc, fc = fcPcd + 2(cp – PlcPcd.aCp[i])
            //    If pcd.fc.fCompressed == 1, fc /= 2, fcPcd /= 2

            // Build PlcPcd.aCp and PlcPcd.aPcd from pieceTable
            if (cp >= fib.CcpText)
                return (int)fib.CcpText - 1;
            var pieces = pieceTable.Pieces;
            int i = -1;
            for (int idx = 0; idx < pieces.Count; idx++)
            {
                if (pieces[idx].CpStart <= cp && cp < pieces[idx].CpEnd)
                {
                    i = idx;
                    break;
                }
            }
            if (i == -1)
                return (int)fib.CcpText - 1;

            var pcd = pieces[i];
            int plcCp = pcd.CpStart;
            int fcPcd = (int)pcd.FcStart;
            int fc = fcPcd + 2 * (cp - plcCp);
            if (!pcd.IsUnicode)
            {
                fc /= 2;
                fcPcd /= 2;
            }

            // Step 4: Read PlcBtePapx from Table Stream
            if (fib.FcPlcfbtePapx <= 0 || fib.LcbPlcfbtePapx <= 0 || tableStream == null)
                throw new InvalidOperationException("No PlcBtePapx available");

            byte[] plcbtePapx = new byte[fib.LcbPlcfbtePapx];
            Array.Copy(tableStream, fib.FcPlcfbtePapx, plcbtePapx, 0, fib.LcbPlcfbtePapx);

            // PlcBtePapx: aFc[] (int32), aPnBtePapx[] (uint16), last aFc is end
            int nPapx = (int)((fib.LcbPlcfbtePapx - 4) / 6);
            int[] aFc = new int[nPapx + 1];
            ushort[] aPnBtePapx = new ushort[nPapx];
            using (var ms = new System.IO.MemoryStream(plcbtePapx))
            using (var br = new System.IO.BinaryReader(ms))
            {
                for (int j = 0; j < nPapx + 1; j++)
                    aFc[j] = br.ReadInt32();
                for (int j = 0; j < nPapx; j++)
                    aPnBtePapx[j] = br.ReadUInt16();
            }

            int fcLast = aFc[^1];
            if (fcLast <= fc)
            {
                if (fcLast < fcPcd)
                {
                    // Step 8: If PlcPcd.aCp[i] is 0, return 0
                    if (plcCp == 0)
                        return 0;
                    // Step 9: Set cp = PlcPcd.aCp[i], i = i-1, repeat
                    return FindParagraphStartCp(plcCp, fib, pieceTable, tableStream, wordDocStream);
                }
                fc = fcLast;
                if (!pcd.IsUnicode)
                    fcLast /= 2;
                // Step 7: If fcFirst > fcPcd, compute dfc and return paragraph start cp
                int dfc = fcLast - fcPcd;
                if (pcd.IsUnicode)
                    dfc /= 2;
                int paraStartCp = plcCp + dfc;
                return paraStartCp;
            }

            // Step 5: Find largest j such that aFc[j] <= fc
            int jIdx = 0;
            for (int j = 0; j < aFc.Length; j++)
            {
                if (aFc[j] <= fc)
                    jIdx = j;
                else
                    break;
            }
            ushort pn = aPnBtePapx[jIdx];

            // Step 6: Read PapxFkp at offset pn*512 in WordDocument stream
            long papxFkpOffset = pn * 512L;
            wordDocStream.Seek(papxFkpOffset, System.IO.SeekOrigin.Begin);
            byte[] papxFkp = new byte[512];
            int read = 0;
            if (papxFkpOffset < wordDocStream.Length)
            {
                int toRead = (int)Math.Min(512, wordDocStream.Length - papxFkpOffset);
                read = wordDocStream.Read(papxFkp, 0, toRead);
            }
            // If less than 512 bytes, pad the rest with zeros (Word does this)
            if (read < 512)
            {
                for (int z = read; z < 512; z++)
                    papxFkp[z] = 0;
            }

            // PapxFkp: first n+1 int32 rgfc, then n Papx, then crun byte at 511
            int crun = papxFkp[511];
            // If the PapxFkp block is empty (read==0), crun will be 0 and the block is invalid.
            if (read == 0 || crun == 0)
            {
                return (int)fib.CcpText - 1;
            }
            int[] rgfc = new int[crun + 1];
            for (int k = 0; k < crun + 1; k++)
                rgfc[k] = BitConverter.ToInt32(papxFkp, k * 4);

            // Handle empty or invalid PapxFkp block
            if (rgfc.Length == 0 || rgfc[^1] == 0)
            {
                return (int)fib.CcpText - 1;
            }

            // Step 6: Find largest k such that rgfc[k] <= fc
            int kIdx = 0;
            for (int k = 0; k < rgfc.Length; k++)
            {
                if (rgfc[k] <= fc)
                    kIdx = k;
                else
                    break;
            }
            if (rgfc[^1] <= fc)
            {
                // If fc is beyond the last rgfc, clamp to end of main text
                return (int)fib.CcpText - 1;
            }

            int fcFirst = rgfc[kIdx];

            // Step 7: If fcFirst > fcPcd, compute dfc and return paragraph start cp
            if (fcFirst > fcPcd)
            {
                int dfc = fcFirst - fcPcd;
                if (pcd.IsUnicode)
                    dfc /= 2;
                int paraStartCp = plcCp + dfc;
                return paraStartCp;
            }
            // Step 8: If PlcPcd.aCp[i] is 0, return 0
            if (plcCp == 0)
                return 0;
            // Step 9: Set cp = PlcPcd.aCp[i], i = i-1, repeat
            return FindParagraphStartCp(plcCp, fib, pieceTable, tableStream, wordDocStream);
        }

        // Implements [MS-DOC] 2.4.2 paragraph boundary algorithm
        // Returns the character position of the last character in the paragraph containing cp
        private int FindParagraphEndCp(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream)
        {
            // Implements [MS-DOC] 2.4.2 for finding paragraph end
            var pieces = pieceTable.Pieces;
            int i = -1;
            for (int idx = 0; idx < pieces.Count; idx++)
            {
                if (pieces[idx].CpStart <= cp && cp < pieces[idx].CpEnd)
                {
                    i = idx;
                    break;
                }
            }
            if (i == -1)
                throw new ArgumentException("Invalid cp: not found in any piece");

            var pcd = pieces[i];
            int plcCp = pcd.CpStart;
            int fcPcd = (int)pcd.FcStart;
            int fc = fcPcd + 2 * (cp - plcCp);
            int fcMac = fcPcd + 2 * (pieces[i].CpEnd - plcCp);
            if (!pcd.IsUnicode)
            {
                fc /= 2;
                fcPcd /= 2;
                fcMac /= 2;
            }

            // Step 4: Read PlcBtePapx from Table Stream
            if (fib.FcPlcfbtePapx <= 0 || fib.LcbPlcfbtePapx <= 0 || tableStream == null)
                throw new InvalidOperationException("No PlcBtePapx available");

            byte[] plcbtePapx = new byte[fib.LcbPlcfbtePapx];
            Array.Copy(tableStream, fib.FcPlcfbtePapx, plcbtePapx, 0, fib.LcbPlcfbtePapx);

            int nPapx = (int)((fib.LcbPlcfbtePapx - 4) / 6);
            int[] aFc = new int[nPapx + 1];
            ushort[] aPnBtePapx = new ushort[nPapx];
            using (var ms = new System.IO.MemoryStream(plcbtePapx))
            using (var br = new System.IO.BinaryReader(ms))
            {
                for (int j = 0; j < nPapx + 1; j++)
                    aFc[j] = br.ReadInt32();
                for (int j = 0; j < nPapx; j++)
                    aPnBtePapx[j] = br.ReadUInt16();
            }

            // Step 4: Find largest j such that aFc[j] <= fc
            int jIdx = 0;
            for (int j = 0; j < aFc.Length; j++)
            {
                if (aFc[j] <= fc)
                    jIdx = j;
                else
                    break;
            }
            ushort pn = aPnBtePapx[jIdx];

            // Step 5: Read PapxFkp at offset pn*512 in WordDocument stream
            long papxFkpOffset = pn * 512L;
            wordDocStream.Seek(papxFkpOffset, System.IO.SeekOrigin.Begin);
            byte[] papxFkp = new byte[512];
            int read = 0;
            if (papxFkpOffset < wordDocStream.Length)
            {
                int toRead = (int)Math.Min(512, wordDocStream.Length - papxFkpOffset);
                read = wordDocStream.Read(papxFkp, 0, toRead);
            }
            // If less than 512 bytes, pad the rest with zeros (Word does this) 
            if (read < 512)
            {
                for (int z = read; z < 512; z++)
                    papxFkp[z] = 0;
            }

            int crun = papxFkp[511];
            // If the PapxFkp block is empty (read==0), crun will be 0 and the block is invalid.
            if (read == 0 || crun == 0)
            {
                return (int)fib.CcpText - 1;
            }
            int[] rgfc = new int[crun + 1];
            for (int k = 0; k < crun + 1; k++)
                rgfc[k] = BitConverter.ToInt32(papxFkp, k * 4);

            // Handle empty or invalid PapxFkp block
            if (rgfc.Length == 0 || rgfc[^1] == 0)
            {
                return (int)fib.CcpText - 1;
            }

            // Step 5: Find largest k such that rgfc[k] <= fc
            int kIdx = 0;
            for (int k = 0; k < rgfc.Length; k++)
            {
                if (rgfc[k] <= fc)
                    kIdx = k;
                else
                    break;
            }
            if (rgfc[^1] <= fc)
            {
                // If fc is beyond the last rgfc, clamp to end of main text
                return (int)fib.CcpText - 1;
            }

            int fcLim = rgfc[kIdx + 1];

            // Step 6: If fcLim <= fcMac, compute dfc and return paragraph end cp
            if (fcLim <= fcMac)
            {
                int dfc = fcLim - fcPcd;
                if (pcd.IsUnicode)
                    dfc /= 2;
                int paraEndCp = plcCp + dfc - 1;
                return paraEndCp;
            }
            // Step 7: Set cp = PlcPcd.aCp[i+1], i = i+1, repeat
            int nextCp = pieces[i].CpEnd;
            int nextIdx = i + 1;
            if (nextIdx >= pieces.Count)
                return pieces[^1].CpEnd - 1;
            return FindParagraphEndCp(nextCp, fib, pieceTable, tableStream, wordDocStream);
        }


        public WordDocumentParser(CompoundFileBinaryFormatParser cfbfParser, ILogger logger)
        {
            _cfbfParser = cfbfParser;
            _logger = logger;
            _documentModel = new Core.DocumentModel();
        }

        private Core.DocumentModel _documentModel;

        private static bool IsAlphaNumeric(byte b)
        {
            char c = (char)b;
            return char.IsLetterOrDigit(c);
        }

        public void ParseDocument(string? password = null)
        {
            // Ensure header is parsed before reading directory entries
            _cfbfParser.ParseHeader();

            // Parse directory entries
            var entries = _cfbfParser.ParseDirectoryEntries();

            // Locate WordDocument stream
            var wordDocEntry = entries.Find(e =>
                e.Name.Contains("WordDocument", StringComparison.OrdinalIgnoreCase));

            if (wordDocEntry == null)
            {
                _logger.LogError("Required WordDocument stream not found in CFBF file. Available directory entries:");
                foreach (var entry in entries)
                {
                    _logger.LogInfo($"- {entry.Name} (Type: {entry.EntryType}, Size: {entry.StreamSize})");
                }
                throw new InvalidDataException("Required WordDocument stream not found in CFBF file.");
            }
            _logger.LogInfo($"[DEBUG] WordDocument Entry: Name='{wordDocEntry.Name}', StartingSectorLocation={wordDocEntry.StartingSectorLocation}, StreamSize={wordDocEntry.StreamSize}");

            // Read stream data
            var wordDocStream = _cfbfParser.ReadStream(wordDocEntry);

            // Log the first 32 bytes of the WordDocument stream for debugging
            _logger.LogInfo($"[DEBUG] WordDocument stream length: {wordDocStream.Length}");
            if (wordDocStream.Length >= 32)
            {
                var headerBytes = new byte[32];
                Array.Copy(wordDocStream, 0, headerBytes, 0, 32);
                _logger.LogInfo($"[DEBUG] WordDocument stream header: {BitConverter.ToString(headerBytes).Replace("-", " ")}");
            }


            // Parse FIB early to detect Word95 before Table stream check
            var fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);
            _logger.LogInfo($"[DEBUG] Parsed FIB. nFib: {fib.NFib}, fComplex: {fib.FComplex}, fEncrypted: {fib.FEncrypted}");


            // Decrypt Word95 documents if necessary
            if (fib.NFib == 0x0065 && fib.FEncrypted)
            {
                if (string.IsNullOrEmpty(password))
                    throw new InvalidDataException("Password required for Word95 decryption.");

                wordDocStream = WvWareNet.Core.Decryptor95.Decrypt(wordDocStream, password, fib.LKey);
                fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);
            }

            var streamName = fib.FWhichTblStm
                ? "1Table"
                : "0Table";

            var tableEntry = entries
                .FirstOrDefault(e =>
                    e.Name.Equals(streamName, StringComparison.OrdinalIgnoreCase));


            // Word95 files: 100=Word6, 101=Word95, 104=Word97 but some Word95 files use 104
            bool isWord95 = fib.NFib == 100 || fib.NFib == 101 || fib.NFib == 104;

            if (tableEntry == null)
            {
                if (isWord95)
                {
                    // For Word95 files, try to proceed without Table stream if not found
                    _logger.LogWarning($"Table stream not found in Word95/Word6 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                }
                else if (fib.NFib == 53200) // Special case for Word95 test file
                {
                    _logger.LogWarning($"Table stream not found in Word95 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                }
                else
                {
                    _logger.LogError($"Table stream not found, and not a recognized Word95/Word6 document (NFib={fib.NFib})");
                    throw new InvalidDataException("Required Table stream not found in CFBF file.");
                }
            }

            _logger.LogInfo($"[DEBUG] Found WordDocument stream: {wordDocEntry.Name}");
            byte[]? tableStream = null;

            if (tableEntry != null)
            {
                _logger.LogInfo($"[DEBUG] Found Table stream: {tableEntry.Name}");
                tableStream = _cfbfParser.ReadStream(tableEntry);
            }

            _documentModel.FileInfo = fib;

            if ((fib.FEncrypted || fib.FCrypto)) // && !isWord95)
                throw new NotSupportedException("Encrypted Word documents are not supported.");

            if (fib.FibVersion == null)
                _logger.LogWarning($"Unknown Word version NFib={fib.NFib}");
            else
                _logger.LogInfo($"Detected Word version: {fib.FibVersion}");

            var pieceTable = WvWareNet.Core.PieceTable.CreateFromStreams(_logger, fib, tableStream, wordDocStream);

            // Extract CHPX (character formatting) data from Table stream using PLCFCHPX
            // PLCFCHPX location is in FIB: FcPlcfbteChpx, LcbPlcfbteChpx
            var chpxList = new List<byte[]>();
            if (tableStream != null && fib.FcPlcfbteChpx > 0 && fib.LcbPlcfbteChpx > 0 && fib.FcPlcfbteChpx + fib.LcbPlcfbteChpx <= tableStream.Length)
            {
                // Parse PLCFCHPX (Piecewise Linear Control File for CHPX)
                byte[] plcfChpx = new byte[fib.LcbPlcfbteChpx];
                Array.Copy(tableStream, fib.FcPlcfbteChpx, plcfChpx, 0, fib.LcbPlcfbteChpx);

                // Each entry: [CP][CHPX offset], last CP is end
                int pieceCount = pieceTable.Pieces.Count;
                using var plcfStream = new System.IO.MemoryStream(plcfChpx);
                using var plcfReader = new System.IO.BinaryReader(plcfStream);

                int[] cpArray = new int[pieceCount + 1];
                for (int i = 0; i < pieceCount + 1; i++)
                    cpArray[i] = plcfReader.ReadInt32();

                int[] chpxOffsetArray = new int[pieceCount];
                for (int i = 0; i < pieceCount; i++)
                    chpxOffsetArray[i] = plcfReader.ReadInt16();

                // CHPX data follows immediately after the arrays
                long chpxDataStart = plcfStream.Position;
                for (int i = 0; i < pieceCount; i++)
                {
                    plcfStream.Position = chpxDataStart + chpxOffsetArray[i];
                    // CHPX is a variable-length structure, but for now, read a fixed size or until next offset
                    int chpxLen = (i < pieceCount - 1)
                        ? chpxOffsetArray[i + 1] - chpxOffsetArray[i]
                        : (int)(plcfStream.Length - (chpxDataStart + chpxOffsetArray[i]));
                    if (chpxLen > 0)
                        chpxList.Add(plcfReader.ReadBytes(chpxLen));
                    else
                        chpxList.Add(null);
                }
            }
            else
            {
                // Fallback: assign null CHPX to all pieces
                for (int i = 0; i < pieceTable.Pieces.Count; i++)
                    chpxList.Add(null);
            }
            pieceTable.AssignChpxToPieces(chpxList);

            // Parse stylesheet
            WvWareNet.Core.Stylesheet stylesheet = new WvWareNet.Core.Stylesheet();
            _logger.LogInfo($"[DEBUG] Stylesheet info: FcStshf={fib.FcStshf}, LcbStshf={fib.LcbStshf}, tableStream.Length={tableStream?.Length ?? 0}");

            if (tableStream != null && fib.FcStshf > 0 && fib.LcbStshf > 0 && fib.FcStshf + fib.LcbStshf <= tableStream.Length)
            {
                byte[] stshData = new byte[fib.LcbStshf];
                Array.Copy(tableStream, fib.FcStshf, stshData, 0, fib.LcbStshf);
                _logger.LogInfo($"[DEBUG] Extracted stylesheet data of {stshData.Length} bytes from offset {fib.FcStshf}");

                // Log first 32 bytes of stylesheet data
                if (stshData.Length >= 32)
                {
                    string hexStr = BitConverter.ToString(stshData, 0, 32);
                    _logger.LogInfo($"[DEBUG] First 32 bytes of stylesheet: {hexStr}");
                }

                stylesheet = WvWareNet.Core.Stylesheet.Parse(stshData);
                _logger.LogInfo($"[DEBUG] Parsed stylesheet with {stylesheet.Styles.Count} styles");

                // Log the first few styles for debugging
                for (int i = 0; i < Math.Min(10, stylesheet.Styles.Count); i++)
                {
                    var style = stylesheet.Styles[i];
                    _logger.LogInfo($"[DEBUG] Style {style.Index}: '{style.Name}'");
                }
            }
            else
            {
                _logger.LogWarning($"No stylesheet found via FIB - FcStshf={fib.FcStshf}, LcbStshf={fib.LcbStshf}");

                // Create basic default stylesheet with standard built-in styles
                stylesheet = new WvWareNet.Core.Stylesheet();
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 0, Name = "Normal" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 1, Name = "heading 1" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 2, Name = "heading 2" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 3, Name = "heading 3" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 4, Name = "heading 4" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 5, Name = "heading 5" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 6, Name = "heading 6" });
                _logger.LogInfo("[DEBUG] Using default built-in styles as fallback");
            }

            // Extract text using PieceTable and populate DocumentModel with sections and paragraphs
            _documentModel = new WvWareNet.Core.DocumentModel();
            _documentModel.FileInfo = fib; // Store FIB in the document model
            _documentModel.Stylesheet = stylesheet; // Store stylesheet in the document model

            // For now, create a single default section.
            // In a more complete implementation, sections would be parsed from the document structure.
            var defaultSection = new WvWareNet.Core.Section();
            _documentModel.Sections.Add(defaultSection);

            using var wordDocMs = new System.IO.MemoryStream(wordDocStream);

            // Log character counts for debugging
            _logger.LogInfo($"[DEBUG] Character counts - Text: {fib.CcpText}, Footnotes: {fib.CcpFtn}, Headers: {fib.CcpHdr}");

            // --- Paragraph boundary detection using [MS-DOC] 2.4.2 algorithm ---
            // Use new FindParagraphStartCp and FindParagraphEndCp methods
            if (tableStream != null && fib.FcPlcfbtePapx > 0 && fib.LcbPlcfbtePapx > 0 && fib.FcPlcfbtePapx + fib.LcbPlcfbtePapx <= tableStream.Length)
            {
                int cp = 0;
                int paraIdx = 0;
                while (cp < fib.CcpText)
                {
                    int paraStart = FindParagraphStartCp(cp, fib, pieceTable, tableStream, wordDocMs);
                    int paraEnd = FindParagraphEndCp(cp, fib, pieceTable, tableStream, wordDocMs);
                    if (paraEnd < paraStart) break;
                    if (paraEnd >= fib.CcpText) paraEnd = (int)fib.CcpText - 1;

                    // Find all pieces that overlap this paragraph
                    var paraPieces = new List<int>();
                    for (int p = 0; p < pieceTable.Pieces.Count; p++)
                    {
                        var piece = pieceTable.Pieces[p];
                        if (piece.CpStart < paraEnd + 1 && piece.CpEnd > paraStart)
                            paraPieces.Add(p);
                    }

                    var paragraph = new WvWareNet.Core.Paragraph();
                    // Style detection from PAPX is not implemented here; can be added if needed
                    paragraph.StyleIndex = 0;
                    paragraph.Style = stylesheet.GetStyleName(0);

                    _logger.LogInfo($"[DEBUG] Paragraph {paraIdx}: CP {paraStart}-{paraEnd}, StyleIndex=0, StyleName='{paragraph.Style}'");

                    foreach (var pIdx in paraPieces)
                    {
                        string text = pieceTable.GetTextForPiece(pIdx, wordDocMs);
                        var chpx = pieceTable.Pieces[pIdx].Chpx;
                        var charProps = new WvWareNet.Core.CharacterProperties();
                        if (chpx != null && chpx.Length > 0)
                        {
                            for (int b = 0; b < chpx.Length - 2; b++)
                            {
                                byte sprm = chpx[b];
                                byte val = chpx[b + 2];
                                switch (sprm)
                                {
                                    case 0x08: charProps.IsBold = val != 0; break;
                                    case 0x09: charProps.IsItalic = val != 0; break;
                                    case 0x0A: charProps.IsStrikeThrough = val != 0; break;
                                    case 0x0D: charProps.IsSmallCaps = val != 0; break;
                                    case 0x0E: charProps.IsAllCaps = val != 0; break;
                                    case 0x0F: charProps.IsHidden = val != 0; break;
                                    case 0x18: charProps.IsUnderlined = val != 0; break;
                                    case 0x2A: charProps.FontSize = val; break;
                                    case 0x2B: break;
                                    case 0x2D: charProps.IsUnderlined = val != 0; break;
                                    case 0x2F: charProps.FtcaAscii = val; break;
                                    case 0x30: charProps.LanguageId = val; break;
                                    case 0x31: charProps.CharacterSpacing = (short)val; break;
                                    case 0x32: charProps.CharacterScaling = val; break;
                                }
                            }
                        }
                        var run = new WvWareNet.Core.Run { Text = text, Properties = charProps };
                        paragraph.Runs.Add(run);
                    }

                    defaultSection.Paragraphs.Add(paragraph);
                    paraIdx++;
                    cp = paraEnd + 1;
                }
            }
            else
            {
                // Fallback: old logic using newlines
                var currentParagraph = new WvWareNet.Core.Paragraph();
                defaultSection.Paragraphs.Add(currentParagraph);

                for (int i = 0; i < pieceTable.Pieces.Count; i++)
                {
                    string text = pieceTable.GetTextForPiece(i, wordDocMs);
                    _logger.LogInfo($"[DEBUG] Extracted text from piece {i}: \"{text.Replace("\r", "\\r").Replace("\n", "\\n")}\"");

                    var chpx = pieceTable.Pieces[i].Chpx;
                    var charProps = new WvWareNet.Core.CharacterProperties();
                    if (chpx != null && chpx.Length > 0)
                    {
                        for (int b = 0; b < chpx.Length - 2; b++)
                        {
                            byte sprm = chpx[b];
                            byte val = chpx[b + 2];
                            switch (sprm)
                            {
                                case 0x08: charProps.IsBold = val != 0; break;
                                case 0x09: charProps.IsItalic = val != 0; break;
                                case 0x0A: charProps.IsStrikeThrough = val != 0; break;
                                case 0x0D: charProps.IsSmallCaps = val != 0; break;
                                case 0x0E: charProps.IsAllCaps = val != 0; break;
                                case 0x0F: charProps.IsHidden = val != 0; break;
                                case 0x18: charProps.IsUnderlined = val != 0; break;
                                case 0x2A: charProps.FontSize = val; break;
                                case 0x2B: break;
                                case 0x2D: charProps.IsUnderlined = val != 0; break;
                                case 0x2F: charProps.FtcaAscii = val; break;
                                case 0x30: charProps.LanguageId = val; break;
                                case 0x31: charProps.CharacterSpacing = (short)val; break;
                                case 0x32: charProps.CharacterScaling = val; break;
                            }
                        }
                    }

                    var run = new WvWareNet.Core.Run { Text = text, Properties = charProps };
                    currentParagraph.Runs.Add(run);

                    if (text.EndsWith("\r\n") || text.EndsWith("\n") || text.EndsWith("\r"))
                    {
                        currentParagraph = new WvWareNet.Core.Paragraph();
                        defaultSection.Paragraphs.Add(currentParagraph);
                    }
                }

                if (defaultSection.Paragraphs.Count > 0 && defaultSection.Paragraphs[^1].Runs.Count == 0)
                {
                    defaultSection.Paragraphs.RemoveAt(defaultSection.Paragraphs.Count - 1);
                }
            }

            // --- Header/Footer/Footnote Extraction ---
            // Parse PLCF for headers/footers if available
            if (tableStream != null && fib.FcPlcfhdd > 0 && fib.LcbPlcfhdd > 0 && fib.FcPlcfhdd + fib.LcbPlcfhdd <= tableStream.Length)
            {
                byte[] plcfhdd = new byte[fib.LcbPlcfhdd];
                Array.Copy(tableStream, fib.FcPlcfhdd, plcfhdd, 0, fib.LcbPlcfhdd);

                // Each entry: [CP][FC], last CP is end
                int entryCount = (plcfhdd.Length - 4) / 8; // 4 bytes CP, 4 bytes FC per entry
                using var plcfStream = new System.IO.MemoryStream(plcfhdd);
                using var plcfReader = new System.IO.BinaryReader(plcfStream);

                for (int i = 0; i < entryCount; i++)
                {
                    int cp = plcfReader.ReadInt32();
                    int fc = plcfReader.ReadInt32();

                    // Extract header/footer text from WordDocument stream at FC
                    string headerText = "";
                    if (fc > 0 && fc < wordDocStream.Length)
                    {
                        using var ms = new System.IO.MemoryStream(wordDocStream);
                        ms.Position = fc;
                        using var reader = new System.IO.BinaryReader(ms);
                        // Read up to 512 bytes or until null terminator (simple heuristic)
                        var bytes = reader.ReadBytes(512);
                        int len = Array.IndexOf(bytes, (byte)0);
                        if (len < 0) len = bytes.Length;
                        headerText = System.Text.Encoding.Default.GetString(bytes, 0, len);
                        // Try Windows-1251 if Default decoding looks garbled (heuristic: many '?')
                        if (headerText.Count(c => c == '?') > headerText.Length / 4)
                        {
                            try
                            {
                                headerText = System.Text.Encoding.GetEncoding(1251).GetString(bytes, 0, len);
                                _logger.LogInfo("[Cyrillic] Header decoded with Windows-1251");
                            }
                            catch { }
                        }
                    }

                    _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter
                    {
                        Type = WvWareNet.Core.HeaderFooterType.Default,
                        Paragraphs = { new WvWareNet.Core.Paragraph { Runs = { new WvWareNet.Core.Run { Text = headerText } } } }
                    });
                }
            }
            else
            {
                _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            }

            // --- Text Box Extraction ---
            if (tableStream != null && fib.FcPlcftxbxTxt > 0 && fib.LcbPlcftxbxTxt > 0 &&
                fib.FcPlcftxbxTxt + fib.LcbPlcftxbxTxt <= tableStream.Length)
            {
                byte[] plcfTxtBx = new byte[fib.LcbPlcftxbxTxt];
                Array.Copy(tableStream, fib.FcPlcftxbxTxt, plcfTxtBx, 0, fib.LcbPlcftxbxTxt);

                // Parse PLCF for text boxes
                using var msPlcf = new System.IO.MemoryStream(plcfTxtBx);
                using var reader = new System.IO.BinaryReader(msPlcf);

                // Read CP and FC arrays
                int nTextBoxes = (plcfTxtBx.Length - 4) / 8;  // Each text box has CP and FC
                int[] cpArray = new int[nTextBoxes + 1];
                int[] fcArray = new int[nTextBoxes];

                for (int i = 0; i <= nTextBoxes; i++)
                    cpArray[i] = reader.ReadInt32();

                for (int i = 0; i < nTextBoxes; i++)
                    fcArray[i] = reader.ReadInt32();

                // Process each text box
                for (int i = 0; i < nTextBoxes; i++)
                {
                    int fcStart = fcArray[i];
                    int fcEnd = (i < nTextBoxes - 1) ? fcArray[i + 1] : (int)wordDocStream.Length;

                    // Extract text box content
                    string text = pieceTable.GetTextForRange(fcStart, fcEnd, wordDocMs);

                    var textBox = new Core.TextBox();
                    var paragraph = new Core.Paragraph();
                    paragraph.Runs.Add(new Core.Run { Text = text });
                    textBox.Paragraphs.Add(paragraph);
                    _documentModel.TextBoxes.Add(textBox);
                }
            }

            // Parse PLCF for footnotes if available
            if (tableStream != null && fib.FcPlcffldFtn > 0 && fib.LcbPlcffldFtn > 0 && fib.FcPlcffldFtn + fib.LcbPlcffldFtn <= tableStream.Length)
            {
                byte[] plcfftn = new byte[fib.LcbPlcffldFtn];
                Array.Copy(tableStream, fib.FcPlcffldFtn, plcfftn, 0, fib.LcbPlcffldFtn);

                int entryCount = (plcfftn.Length - 4) / 8; // 4 bytes CP, 4 bytes FC per entry
                using var plcfStream = new System.IO.MemoryStream(plcfftn);
                using var plcfReader = new System.IO.BinaryReader(plcfStream);

                for (int i = 0; i < entryCount; i++)
                {
                    int cp = plcfReader.ReadInt32();
                    int fc = plcfReader.ReadInt32();

                    // Extract footnote text from WordDocument stream at FC
                    string footnoteText = "";
                    if (fc > 0 && fc < wordDocStream.Length)
                    {
                        using var ms = new System.IO.MemoryStream(wordDocStream);
                        ms.Position = fc;
                        using var reader = new System.IO.BinaryReader(ms);
                        // Read up to 512 bytes or until null terminator (simple heuristic)
                        var bytes = reader.ReadBytes(512);
                        int len = Array.IndexOf(bytes, (byte)0);
                        if (len < 0) len = bytes.Length;
                        footnoteText = System.Text.Encoding.Default.GetString(bytes, 0, len);
                        // Try Windows-1251 if Default decoding looks garbled (heuristic: many '?')
                        if (footnoteText.Count(c => c == '?') > footnoteText.Length / 4)
                        {
                            try
                            {
                                footnoteText = System.Text.Encoding.GetEncoding(1251).GetString(bytes, 0, len);
                                _logger.LogInfo("[Cyrillic] Footnote decoded with Windows-1251");
                            }
                            catch { }
                        }
                    }

                    _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote
                    {
                        ReferenceId = i + 1,
                        Paragraphs = { new WvWareNet.Core.Paragraph { Runs = { new WvWareNet.Core.Run { Text = footnoteText } } } }
                    });
                }
            }
            else
            {
                _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote { ReferenceId = 1 });
            }

            // Add a default footer if not already added
            if (_documentModel.Footers.Count == 0)
            {
                _documentModel.Footers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            }

            // Log document structure for debugging
            _logger.LogInfo($"[STRUCTURE] Document has {_documentModel.Sections.Count} section(s)");
            for (int i = 0; i < _documentModel.Sections.Count; i++)
            {
                var section = _documentModel.Sections[i];
                _logger.LogInfo($"[STRUCTURE] Section {i + 1}: {section.Paragraphs.Count} paragraph(s)");
                for (int j = 0; j < section.Paragraphs.Count; j++)
                {
                    var paragraph = section.Paragraphs[j];
                    _logger.LogInfo($"[STRUCTURE] Paragraph {j + 1}: {paragraph.Runs.Count} run(s)");
                    for (int k = 0; k < paragraph.Runs.Count; k++)
                    {
                        var run = paragraph.Runs[k];
                        _logger.LogInfo($"[STRUCTURE] Run {k + 1}: Length={run.Text?.Length ?? 0}, TextPreview='{(run.Text != null ? run.Text.Substring(0, Math.Min(20, run.Text.Length)) : "null")}'");
                    }
                }
            }

            _logger.LogInfo("[PARSING] Document parsing completed");
        }


        public string ExtractText()
        {
            if (_documentModel == null)
                throw new InvalidOperationException("Document not parsed. Call ParseDocument() first.");

            var textBuilder = new System.Text.StringBuilder();
            int runCount = 0;
            int charCount = 0;

            // Extract text from document body only (ignore headers/footers/notes)
            foreach (var section in _documentModel.Sections)
            {
                foreach (var paragraph in section.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        if (run.Text != null)
                        {
                            textBuilder.Append(run.Text);
                            runCount++;
                            charCount += run.Text.Length;
                        }
                    }
                }
            }

            _logger.LogInfo($"[EXTRACTION] Extracted {runCount} runs with {charCount} characters");
            return textBuilder.ToString();
        }
    }
}

```

```csharp
// PieceTable.cs (That works with the most documents, currently the head version)
using System;
using System.Collections.Generic;
using System.IO;
using WvWareNet.Utilities;

namespace WvWareNet.Core;

public class PieceTable
{
    private readonly ILogger _logger;
    private readonly List<PieceDescriptor> _pieces = new();

    public IReadOnlyList<PieceDescriptor> Pieces => _pieces;

    public PieceTable(ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Factory method to create and parse a PieceTable from the given streams and FIB.
    /// Handles CLX extraction and all fallbacks.
    /// </summary>
    public static PieceTable CreateFromStreams(
        ILogger logger,
        FileInformationBlock fib,
        byte[]? tableStream,
        byte[] wordDocStream)
    {
        var pieceTable = new PieceTable(logger);

        byte[]? clxData = null;
        if (tableStream != null)
        {
            // Word 97 and later use FcClx and LcbClx from FIB
            if (fib.NFib >= 0x0076) // Word 97 (0x0076) and later
            {
                if (fib.LcbClx > 0 && fib.FcClx >= 0 && (fib.FcClx + fib.LcbClx) <= tableStream.Length)
                {
                    logger.LogInfo($"[DEBUG] FIB: FcClx={fib.FcClx}, LcbClx={fib.LcbClx}, tableStream.Length={tableStream.Length}");
                    clxData = new byte[fib.LcbClx];
                    Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);
                    logger.LogInfo($"[DEBUG] CLX data: {BitConverter.ToString(clxData)}");
                }
            }
            else // Word 6/95 - CLX is at the beginning of the Table stream
            {
                // Heuristic: Assume CLX is at the beginning of the table stream for older formats
                // and has a reasonable size (e.g., up to 512 bytes, or the whole stream if smaller)
                int assumedClxLength = Math.Min(tableStream.Length, 512);
                clxData = new byte[assumedClxLength];
                Array.Copy(tableStream, 0, clxData, 0, assumedClxLength);
                logger.LogInfo($"[DEBUG] Assuming CLX at start of Table stream for NFib={fib.NFib}, length={assumedClxLength}");
            }
        }

        if (clxData != null && clxData.Length > 0)
        {
            logger.LogInfo($"[DEBUG] Parsing piece table with CLX data (size: {clxData.Length}), FcMin: {fib.FcMin}, FcMac: {fib.FcMac}, NFib: {fib.NFib}");
            pieceTable.Parse(clxData, fib);

            logger.LogInfo($"[DEBUG] Piece table parsed. Number of pieces: {pieceTable.Pieces.Count}");
            foreach (var piece in pieceTable.Pieces)
            {
                logger.LogInfo($"[DEBUG]   - Piece: CpStart={piece.CpStart}, CpEnd={piece.CpEnd}, FcStart={piece.FcStart}, IsUnicode={piece.IsUnicode}");
            }

            if (pieceTable.Pieces.Count == 1 && pieceTable.Pieces[0].FcStart >= wordDocStream.Length)
            {
                logger.LogWarning("Invalid piece table detected, falling back to FcMin/FcMac range.");
                pieceTable.SetSinglePiece(fib);
            }
        }
        else
        {
            logger.LogWarning("No CLX data found or invalid, treating document as single piece");
            pieceTable.SetSinglePiece(fib);
            logger.LogInfo($"[DEBUG] Fallback single piece created. IsUnicode: {pieceTable.Pieces[0].IsUnicode}");
        }

        return pieceTable;
    }

    /// <summary>
    /// Assigns CHPX (character formatting) data to each piece descriptor.
    /// </summary>
    public void AssignChpxToPieces(List<byte[]> chpxList)
    {
        // Assumes chpxList.Count == _pieces.Count or chpxList.Count == _pieces.Count - 1
        int count = Math.Min(_pieces.Count, chpxList.Count);
        for (int i = 0; i < count; i++)
        {
            _pieces[i].Chpx = chpxList[i];
        }
    }

    public void Parse(byte[] clxData, FileInformationBlock fib)
    {
        _pieces.Clear();
        if (clxData == null || clxData.Length == 0)
        {
            _logger.LogWarning("Attempted to parse empty piece table data");
            return;
        }

        // DEBUG: Print first 32 bytes of clxData and its length
        Console.WriteLine($"[DEBUG] PieceTable.Parse: data.Length={clxData.Length}");
        Console.WriteLine($"[DEBUG] PieceTable.Parse: first 32 bytes: {BitConverter.ToString(clxData, 0, Math.Min(32, clxData.Length))}");
        Console.WriteLine("[DEBUG] PieceTable.Parse: full CLX data:");
        for (int dbg_i = 0; dbg_i < clxData.Length; dbg_i++)
        {
            Console.Write($"{clxData[dbg_i]:X2} ");
            if ((dbg_i + 1) % 16 == 0) Console.WriteLine();
        }
        Console.WriteLine();

        // --- CLX parsing logic ---
        // CLX can be a sequence of [0x01 Prc] and [0x02 PlcPcd] blocks.
        // We want to find the 0x02 block and parse it as the piece table.
        // However, some documents may have different formats or padding.

        int plcPcdOffset = -1;
        int plcPcdLength = -1;

        int i = 0;
        
        // Skip leading zeros (padding)
        while (i < clxData.Length && clxData[i] == 0)
            i++;

        while (i < clxData.Length)
        {
            byte clxType = clxData[i];
            if (clxType == 0x02)
            {
                // This is the Pcdt block. The next four bytes give the
                // length of the PlcPcd structure that follows.
                if (i + 5 > clxData.Length)
                    break;
                plcPcdLength = BitConverter.ToInt32(clxData, i + 1);
                plcPcdOffset = i + 5;
                _logger.LogInfo($"[DEBUG] Found PlcPcd at offset {i}, length {plcPcdLength}");
                break;
            }
            else if (clxType == 0x01)
            {
                // This is a Prc (property modifier) block, skip it
                if (i + 3 > clxData.Length)
                    break;
                ushort prcSize = BitConverter.ToUInt16(clxData, i + 1);
                _logger.LogInfo($"[DEBUG] Skipping Prc block at offset {i}, size {prcSize}");
                i += 1 + 2 + prcSize;
            }
            else
            {
                // For some documents, the piece table may start directly without clxt prefix
                // Try to parse the remainder as a piece table
                _logger.LogInfo($"[DEBUG] No clxt prefix found, trying direct piece table parse from offset {i}");
                plcPcdOffset = i;
                plcPcdLength = clxData.Length - i;
                break;
            }
        }

        if (plcPcdOffset == -1)
        {
            _logger.LogWarning("No PlcPcd (piece table) found in CLX data. Falling back to single piece.");
            SetSinglePiece(fib);
            return;
        }

        // Now parse the PlcPcd as the piece table
        try
        {
            using var stream = new MemoryStream(clxData, plcPcdOffset, plcPcdLength);
            using var reader = new BinaryReader(stream);

            _logger.LogInfo($"[DEBUG] Attempting to parse piece table at offset {plcPcdOffset}, length {plcPcdLength}");
            
            // For debugging, log the first 16 bytes of what we're trying to parse
            if (plcPcdLength >= 16)
            {
                byte[] debugBytes = new byte[16];
                Array.Copy(clxData, plcPcdOffset, debugBytes, 0, 16);
                _logger.LogInfo($"[DEBUG] First 16 bytes of piece table data: {BitConverter.ToString(debugBytes)}");
            }

            // The piece table (PlcPcd) consists of an array of character
            // positions followed by an array of piece descriptors.  
            // Each piece: 4 bytes CP, 8 bytes descriptor (total 12 bytes per piece)
            // However, the structure might be different for some documents
            
            // Try to detect if this looks like valid piece table data
            // Valid CP values should be reasonable (0 to document length)
            // Let's try different interpretations
            
            bool validPieceTable = false;
            int pieceCount = 0;
            
            // Try standard interpretation first
            if (plcPcdLength >= 16) // Minimum for 1 piece (4+4+8)
            {
                int testPieceCount = (plcPcdLength - 4) / 12;
                if (testPieceCount > 0 && testPieceCount < 20) // Reasonable range
                {
                    // Test if the first few CP values look reasonable
                    stream.Position = 0;
                    int cp1 = reader.ReadInt32();
                    int cp2 = reader.ReadInt32();
                    
                    _logger.LogInfo($"[DEBUG] Test CP values: cp1={cp1}, cp2={cp2}");
                    
                    if (cp1 >= 0 && cp1 < 100000 && cp2 >= cp1 && cp2 < 100000)
                    {
                        validPieceTable = true;
                        pieceCount = testPieceCount;
                        _logger.LogInfo($"[DEBUG] Standard piece table format detected, {pieceCount} pieces");
                    }
                }
            }
            
            if (!validPieceTable)
            {
                _logger.LogWarning($"[DEBUG] Piece table data doesn't match expected format, falling back to single piece");
                SetSinglePiece(fib);
                return;
            }

            // Parse the piece table using standard format
            stream.Position = 0;
            var cpArray = new int[pieceCount + 1];
            for (int j = 0; j < pieceCount + 1; j++)
            {
                cpArray[j] = reader.ReadInt32();
                _logger.LogInfo($"[DEBUG] CP[{j}] = {cpArray[j]}");
            }

            for (int j = 0; j < pieceCount; j++)
            {
                uint fcValue = reader.ReadUInt32();
                uint prm = reader.ReadUInt32(); // PRM (property modifier)
                
                _logger.LogInfo($"[DEBUG] Piece {j}: fcValue=0x{fcValue:X8}, prm=0x{prm:X8}");

                // According to MS-DOC spec, the FC value interpretation:
                // - If bit 30 (0x40000000) is SET: 8-bit characters, fc = (fcValue & ~0x40000000) / 2
                // - If bit 30 is CLEAR: 16-bit characters, fc = fcValue & ~0x40000000
                
                bool isCompressed = (fcValue & 0x40000000) != 0;
                bool isUnicode = !isCompressed;
                uint fc;
                
                if (isCompressed)
                {
                    // 8-bit characters: clear the compression bit and divide by 2
                    fc = (fcValue & 0xBFFFFFFF) / 2;
                }
                else
                {
                    // 16-bit characters: just clear any high bits
                    fc = fcValue & 0x3FFFFFFF;
                }

                _logger.LogInfo($"fcValue=0x{fcValue:X8}, isCompressed={isCompressed}, isUnicode={isUnicode}, fc={fc}");

                int cpStart = cpArray[j];
                int cpEnd = cpArray[j + 1];
                int fcStart = (int)fc;
                int fcEnd = fcStart + (isUnicode ? (cpEnd - cpStart) * 2 : (cpEnd - cpStart));

                _logger.LogInfo($"[DEBUG] Piece {j}: CP {cpStart}-{cpEnd}, FC {fcStart}-{fcEnd}, Unicode={isUnicode}");

                var descriptor = new PieceDescriptor
                {
                    FilePosition = fc,
                    IsUnicode = isUnicode,
                    HasFormatting = false,
                    CpStart = cpStart,
                    CpEnd = cpEnd,
                    FcStart = fcStart,
                    FcEnd = fcEnd
                };

                _pieces.Add(descriptor);
            }

            _logger.LogInfo($"Parsed {_pieces.Count} piece descriptors from piece table");
        }
        catch (Exception ex)
        {
            _logger.LogError("Error parsing piece table data", ex);
            _logger.LogWarning("Falling back to single piece due to parsing error");
            SetSinglePiece(fib);
        }
    }

    public string GetTextForPiece(int index, Stream documentStream, int? codePage = null)
    {
        if (index < 0 || index >= _pieces.Count)
            throw new ArgumentOutOfRangeException(nameof(index));

        var piece = _pieces[index];
        return GetTextForRange(piece.FcStart, piece.FcEnd, documentStream, piece.IsUnicode, codePage);
    }

    /// <summary>
    /// Retrieve text for an arbitrary file position range. The range may span
    /// multiple pieces and does not need to align to piece boundaries.
    /// </summary>
    public string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, int? codePage = null)
    {
        if (fcEnd <= fcStart)
            return string.Empty;

        var sb = new System.Text.StringBuilder();

        foreach (var piece in _pieces)
        {
            int start = Math.Max(fcStart, piece.FcStart);
            int end = Math.Min(fcEnd, piece.FcEnd);
            if (start >= end)
                continue;

            sb.Append(GetTextForRange(start, end, documentStream, piece.IsUnicode, codePage));
        }

        return sb.ToString();
    }

    private string GetTextForRange(int fcStart, int fcEnd, Stream documentStream, bool isUnicode, int? codePage = null)
    {
        _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart={fcStart}, fcEnd={fcEnd}, isUnicode={isUnicode}, streamLength={documentStream.Length}, codePage={codePage}");
        int length = fcEnd - fcStart;
        if (length <= 0)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Invalid length {length}, returning empty string");
            return string.Empty;
        }

        if (fcStart >= documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: fcStart {fcStart} >= streamLength {documentStream.Length}, returning empty string");
            return string.Empty;
        }

        // Clamp to stream bounds
        if (fcEnd > documentStream.Length)
        {
            _logger.LogInfo($"[DEBUG] GetTextForRange: Clamping fcEnd from {fcEnd} to {documentStream.Length}");
            fcEnd = (int)documentStream.Length;
            length = fcEnd - fcStart;
        }

        using var reader = new BinaryReader(documentStream, System.Text.Encoding.UTF8, leaveOpen: true);
        documentStream.Seek(fcStart, SeekOrigin.Begin);
        byte[] bytes = reader.ReadBytes(length);
        
        _logger.LogInfo($"[DEBUG] GetTextForRange: Read {bytes.Length} bytes from offset {fcStart}");
        if (bytes.Length > 0)
        {
            string hexPreview = BitConverter.ToString(bytes, 0, Math.Min(32, bytes.Length));
            _logger.LogInfo($"[DEBUG] GetTextForRange: First bytes: {hexPreview}");
        }

        string text;
        if (isUnicode)
        {
            text = System.Text.Encoding.Unicode.GetString(bytes);
            _logger.LogInfo("[DEBUG] GetTextForRange: Decoded as UTF-16LE");
        }
        else
        {
            // Use code page if provided, otherwise default to Windows-1252
            var encoding = codePage.HasValue ? System.Text.Encoding.GetEncoding(codePage.Value) : System.Text.Encoding.GetEncoding(1252);
            text = encoding.GetString(bytes);
            _logger.LogInfo($"[DEBUG] GetTextForRange: Decoded as {(codePage.HasValue ? $"code page {codePage.Value}" : "Windows-1252")}");
        }

        string processedText = ProcessFieldCodes(text);
        _logger.LogInfo($"[DEBUG] GetTextForRange raw text length: {text.Length}, processed length: {processedText.Length}");
        _logger.LogInfo($"[DEBUG] GetTextForRange returning: '{processedText.Substring(0, Math.Min(50, processedText.Length))}'");
        return CleanText(processedText);
    }

    /// <summary>
    /// Process Word field codes and extract only the field results.
    /// Field structure: [0x13][field code][0x14][field result][0x15]
    /// We want to keep only the field result part.
    /// For HYPERLINK fields without field result, extract the URL from the field code.
    /// </summary>
    private static string ProcessFieldCodes(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        int i = 0;
        
        while (i < input.Length)
        {
            char c = input[i];
            
            if (c == (char)0x13) // Field begin (0x13)
            {
                // Find the field separator (0x14) and field end (0x15)
                int separatorPos = -1;
                int endPos = -1;
                int depth = 1; // Track nested fields
                
                for (int j = i + 1; j < input.Length; j++)
                {
                    if (input[j] == (char)0x13) // Nested field begin
                    {
                        depth++;
                    }
                    else if (input[j] == (char)0x15) // Field end
                    {
                        depth--;
                        if (depth == 0)
                        {
                            endPos = j;
                            break;
                        }
                    }
                    else if (input[j] == (char)0x14 && depth == 1 && separatorPos == -1) // Field separator at our level
                    {
                        separatorPos = j;
                    }
                }
                
                if (endPos != -1)
                {
                    if (separatorPos != -1)
                    {
                        // Extract field result (between separator and end)
                        string fieldResult = input.Substring(separatorPos + 1, endPos - separatorPos - 1);
                        sb.Append(fieldResult);
                    }
                    else
                    {
                        // No field separator found, extract field code and check if it's a HYPERLINK
                        string fieldCode = input.Substring(i + 1, endPos - i - 1);
                        
                        // Handle HYPERLINK fields
                        if (fieldCode.StartsWith(" HYPERLINK ", StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract URL from HYPERLINK field
                            string url = fieldCode.Substring(11).Trim(); // Remove " HYPERLINK "
                            
                            // Remove any additional parameters or quotes
                            int spaceIndex = url.IndexOf(' ');
                            if (spaceIndex > 0)
                            {
                                url = url.Substring(0, spaceIndex);
                            }
                            
                            // Remove quotes if present
                            url = url.Trim('"');
                            
                            sb.Append(url);
                        }
                        else
                        {
                            // For other field types without separator, skip the field
                        }
                    }
                    // Skip to after the field end
                    i = endPos + 1;
                }
                else
                {
                    // Malformed field, just append the character and continue
                    sb.Append(c);
                    i++;
                }
            }
            else if (c == (char)0x14 || c == (char)0x15) // Standalone field separators/ends (shouldn't happen in well-formed text)
            {
                // Skip these control characters
                i++;
            }
            else
            {
                // Regular character, append it
                sb.Append(c);
                i++;
            }
        }
        
        return sb.ToString();
    }

    private static string CleanText(string input)
    {
        var sb = new System.Text.StringBuilder(input.Length);
        foreach (char c in input)
        {
            // Handle special Word control characters
            if (c == (char)0x07) // Word tab character (0x07) -> convert to standard tab
            {
                sb.Append('\t');
            }
            else if (c == (char)0x0B) // Word line break (0x0B) -> convert to newline
            {
                sb.Append('\n');
            }
            // Allow a much wider range of characters - be more permissive
            else if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c) ||
                char.IsSymbol(c) ||  
                // TODO: This seems redundany, review it
                c == '\r' || c == '\n' || c == '\t' || c == '\v' || c == '~' ||
                (c >= 'А' && c <= 'я') || // Cyrillic range
                (c >= 'À' && c <= 'ÿ') || // Latin-1 Supplement
                c == '=' || c == '<' || c == '>' || c == '^' || c == '|' || c == '+' || c == '-') // Explicitly allow missing chars
            {
                sb.Append(c);
            }
            else if (c == '\0' || c < ' ') // Replace null chars and control chars
            {
                // Skip these characters
            }
            else
            {
                // For debugging: log any character that's being filtered out
                Console.WriteLine($"[DEBUG] CleanText filtering out character: '{c}' (U+{(int)c:X4})");
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Attempt to determine if the pieces in a Word95 document contain
    /// 16-bit text. This mirrors the heuristic used by wvGuess16bit in the
    /// original wvWare project.
    /// </summary>
    /// <param name="fcValues">File positions for each piece.</param>
    /// <param name="cpArray">Character position array from the piece table.</param>
    private static bool GuessPiecesAre16Bit(uint[] fcValues, int[] cpArray)
    {
        var tuples = new List<(uint Fc, uint Offset)>();
        for (int i = 0; i < fcValues.Length; i++)
        {
            uint offset = (uint)(cpArray[i + 1] - cpArray[i]) * 2u;
            tuples.Add((fcValues[i], offset));
        }

        tuples.Sort((a, b) => a.Fc.CompareTo(b.Fc));

        for (int i = 0; i < tuples.Count - 1; i++)
        {
            if (tuples[i].Fc + tuples[i].Offset > tuples[i + 1].Fc)
                return false; // Overlap means 8-bit text
        }

        return true; // No overlap detected -> assume 16-bit
    }

    /// <summary>
    /// Replace the current table with a single Unicode piece using the
    /// supplied file positions. Used as a fallback when the piece table is
    /// corrupt or not present.
    /// </summary>
    public void SetSinglePiece(FileInformationBlock fib)
    {
        _pieces.Clear();
        // Always use 8-bit encoding for fallback, as in the working version.
        bool isUnicode = false;
        _pieces.Add(new PieceDescriptor
        {
            FilePosition = fib.FcMin,
            IsUnicode = isUnicode,
            HasFormatting = false,
            CpStart = 0,
            CpEnd = (int)(fib.FcMac - fib.FcMin),
            FcStart = (int)fib.FcMin,
            FcEnd = (int)fib.FcMac
        });
    }
}

```

```csharp
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
    //public int FcEnd { get; set; }
    public int FcEnd => FcStart + (IsUnicode ? (CpEnd - CpStart) * 2 : (CpEnd - CpStart));


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

    public uint RawFc { get; set; } // The original, unmodified FC value from the file
    public uint Prm { get; set; }

}
```

```c
/* clx.c */


#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "wv.h"

void
wvReleaseCLX (CLX * clx)
{
    U16 i;
    for (i = 0; i < clx->grpprl_count; i++)
	wvFree (clx->grpprl[i]);
    wvFree (clx->grpprl);
    wvFree (clx->cbGrpprl);
    wvReleasePCD_PLCF (clx->pcd, clx->pos);
}

void
wvBuildCLXForSimple6 (CLX * clx, FIB * fib)
{
    wvInitCLX (clx);
    clx->nopcd = 1;;

    clx->pcd = (PCD *) wvMalloc (clx->nopcd * sizeof (PCD));
    clx->pos = (U32 *) wvMalloc ((clx->nopcd + 1) * sizeof (U32));

    clx->pos[0] = 0;
    clx->pos[1] = fib->ccpText;

    wvInitPCD (&(clx->pcd[0]));
    clx->pcd[0].fc = fib->fcMin;

    /* reverse the special encoding thing they do for word97 
       if we are using the usual 8 bit chars */

    if (fib->fExtChar == 0)
      {
	  clx->pcd[0].fc *= 2;
	  clx->pcd[0].fc |= 0x40000000UL;
      }

    clx->pcd[0].prm.fComplex = 0;
    clx->pcd[0].prm.para.var1.isprm = 0;
    /*
       these set the ones that *I* use correctly, but may break for other wv
       users, though i doubt it, im just marking a possible firepoint for the
       future
     */
}

/*
The complex part of a file (CLX) is composed of a number of variable-sized
blocks of data. Recorded first are any grpprls that may be referenced by the
plcfpcd (if the plcfpcd has no grpprl references, no grpprls will be
recorded) followed by the plcfpcd. Each block in the complex part is
prefaced by a clxt (clx type), which is a 1-byte code, either 1 (meaning the
block contains a grpprl) or 2 (meaning this is the plcfpcd). A clxtGrpprl
(1) is followed by a 2-byte cb which is the count of bytes of the grpprl. A
clxtPlcfpcd (2) is followed by a 4-byte lcb which is the count of bytes of
the piece table. A full saved file will have no clxtGrpprl's.
*/
void
wvGetCLX (wvVersion ver, CLX * clx, U32 offset, U32 len, U8 fExtChar,
	  wvStream * fd)
{
    U8 clxt;
    U16 cb;
    U32 lcb, i, j = 0;

    wvTrace (("offset %x len %d\n", offset, len));
    wvStream_goto (fd, offset);

    wvInitCLX (clx);

    while (j < len)
      {
	  clxt = read_8ubit (fd);
	  j++;
	  if (clxt == 1)
	    {
		cb = read_16ubit (fd);
		j += 2;
		clx->grpprl_count++;
		clx->cbGrpprl =
		    (U16 *) realloc (clx->cbGrpprl,
				     sizeof (U16) * clx->grpprl_count);
		clx->cbGrpprl[clx->grpprl_count - 1] = cb;
		clx->grpprl =
		    (U8 **) realloc (clx->grpprl,
				     sizeof (U8 *) * (clx->grpprl_count));
		clx->grpprl[clx->grpprl_count - 1] = (U8 *) wvMalloc (cb);
		for (i = 0; i < cb; i++)
		    clx->grpprl[clx->grpprl_count - 1][i] = read_8ubit (fd);
		j += i;
	    }
	  else if (clxt == 2)
	    {
		if (ver == WORD8)
		  {
		      lcb = read_32ubit (fd);
		      j += 4;
		  }
		else
		  {
		      wvTrace (("Here so far\n"));
#if 0
		      lcb = read_16ubit (fd);	/* word 6 only has two bytes here */
		      j += 2;
#endif

		      lcb = read_32ubit (fd);	/* word 6 specs appeared to have lied ! */
		      j += 4;
		  }
		wvGetPCD_PLCF (&clx->pcd, &clx->pos, &clx->nopcd,
			       wvStream_tell (fd), lcb, fd);
		j += lcb;

		if (ver <= WORD7)	/* MV 28.8.2000 Appears to be valid */
		  {
#if 0
		      /* DANGER !!, this is a completely mad attempt to differenciate 
		         between word 95 files that use 16 and 8 bit characters. It may
		         not work, it attempt to err on the side of 8 bit characters.
		       */
		      if (!(wvGuess16bit (clx->pcd, clx->pos, clx->nopcd)))
#else
		      /* I think that this is the correct reason for this behaviour */
		      if (fExtChar == 0)
#endif
			  for (i = 0; i < clx->nopcd; i++)
			    {
				clx->pcd[i].fc *= 2;
				clx->pcd[i].fc |= 0x40000000UL;
			    }
		  }
	    }
	  else
	    {
		wvError (("clxt is not 1 or 2, it is %d\n", clxt));
		return;
	    }
      }
}


void
wvInitCLX (CLX * item)
{
    item->pcd = NULL;
    item->pos = NULL;
    item->nopcd = 0;

    item->grpprl_count = 0;
    item->cbGrpprl = NULL;
    item->grpprl = NULL;
}


int
wvGetPieceBoundsFC (U32 * begin, U32 * end, CLX * clx, U32 piececount)
{
    int type;
    if ((piececount + 1) > clx->nopcd)
      {
	  wvTrace (
		   ("piececount is > nopcd, i.e.%d > %d\n", piececount + 1,
		    clx->nopcd));
	  return (-1);
      }
    *begin = wvNormFC (clx->pcd[piececount].fc, &type);

    if (type)
	*end = *begin + (clx->pos[piececount + 1] - clx->pos[piececount]);
    else
	*end = *begin + ((clx->pos[piececount + 1] - clx->pos[piececount]) * 2);

    return (type);
}

int
wvGetPieceBoundsCP (U32 * begin, U32 * end, CLX * clx, U32 piececount)
{
    if ((piececount + 1) > clx->nopcd)
	return (-1);
    *begin = clx->pos[piececount];
    *end = clx->pos[piececount + 1];
    return (0);
}


char *
wvAutoCharset (wvParseStruct * ps)
{
    U16 i = 0;
    int flag;
    char *ret;
    ret = "iso-8859-15";

    /* 
       If any of the pieces use unicode then we have to assume the
       worst and use utf-8
     */
    while (i < ps->clx.nopcd)
      {
	  wvNormFC (ps->clx.pcd[i].fc, &flag);
	  if (flag == 0)
	    {
		ret = "UTF-8";
		break;
	    }
	  i++;
      }

    /* 
       Also if the document fib is not codepage 1252 we also have to 
       assume the worst 
     */
    if (strcmp (ret, "UTF-8"))
      {
	  if (
	      (ps->fib.lid != 0x407) &&
	      (ps->fib.lid != 0x807) &&
	      (ps->fib.lid != 0x409) &&
	      (ps->fib.lid != 0x807) && (ps->fib.lid != 0xC09))
	      ret = "UTF-8";
      }
    return (ret);
}



int
wvQuerySamePiece (U32 fcTest, CLX * clx, U32 piece)
{
    /*
       wvTrace(("Same Piece, %x %x %x\n",fcTest,wvNormFC(clx->pcd[piece].fc,NULL),wvNormFC(clx->pcd[piece+1].fc,NULL)));
       if ( (fcTest >= wvNormFC(clx->pcd[piece].fc,NULL)) && (fcTest < wvNormFC(clx->pcd[piece+1].fc,NULL)) )
     */
    wvTrace (
	     ("Same Piece, %x %x %x\n", fcTest, clx->pcd[piece].fc,
	      wvGetEndFCPiece (piece, clx)));
    if ((fcTest >= wvNormFC (clx->pcd[piece].fc, NULL))
	&& (fcTest < wvGetEndFCPiece (piece, clx)))
	return (1);
    return (0);
}


U32
wvGetPieceFromCP (U32 currentcp, CLX * clx)
{
    U32 i = 0;
    while (i < clx->nopcd)
      {
	  wvTrace (
		   ("i %d: currentcp is %d, clx->pos[i] is %d, clx->pos[i+1] is %d\n",
		    i, currentcp, clx->pos[i], clx->pos[i + 1]));
	  if ((currentcp >= clx->pos[i]) && (currentcp < clx->pos[i + 1]))
	      return (i);
	  i++;
      }
    wvTrace (("cp was not in any piece ! \n", currentcp));
    return (0xffffffffL);
}

U32
wvGetEndFCPiece (U32 piece, CLX * clx)
{
    int flag;
    U32 fc;
    U32 offset = clx->pos[piece + 1] - clx->pos[piece];

    wvTrace (("offset is %x, befc is %x\n", offset, clx->pcd[piece].fc));
    fc = wvNormFC (clx->pcd[piece].fc, &flag);
    wvTrace (("fc is %x, flag %d\n", fc, flag));
    if (flag)
	fc += offset;
    else
	fc += offset * 2;
    wvTrace (("fc is finally %x\n", fc));
    return (fc);
}

/*
1) search for the piece containing the character in the piece table.

2) Then calculate the FC in the file that stores the character from the piece
    table information.
*/
U32
wvConvertCPToFC (U32 currentcp, CLX * clx)
{
    U32 currentfc = 0xffffffffL;
    U32 i = 0;
    int flag;

    while (i < clx->nopcd)
      {
	  if ((currentcp >= clx->pos[i]) && (currentcp < clx->pos[i + 1]))
	    {
		currentfc = wvNormFC (clx->pcd[i].fc, &flag);
		if (flag)
		    currentfc += (currentcp - clx->pos[i]);
		else
		    currentfc += ((currentcp - clx->pos[i]) * 2);
		break;
	    }
	  i++;
      }

    if (currentfc == 0xffffffffL)
      {
	  i--;
	  currentfc = wvNormFC (clx->pcd[i].fc, &flag);
	  if (flag)
	      currentfc += (currentcp - clx->pos[i]);
	  else
	      currentfc += ((currentcp - clx->pos[i]) * 2);
	  wvTrace (("flaky cp to fc conversion underway\n"));
      }

    return (currentfc);
}

struct test {
    U32 fc;
    U32 offset;
};

int
compar (const void *a, const void *b)
{
    struct test *one, *two;
    one = (struct test *) a;
    two = (struct test *) b;

    if (one->fc < two->fc)
	return (-1);
    else if (one->fc == two->fc)
	return (0);
    return (1);
}

/* 
In word 95 files there is no flag attached to each
offset as there is in word 97 to tell you that we are
talking about 16 bit chars, so I attempt here to make
an educated guess based on overlapping offsets to
figure it out, If I had some actual information as
the how word 95 actually stores it it would help.
*/

int
wvGuess16bit (PCD * pcd, U32 * pos, U32 nopcd)
{
    struct test *fcs;
    U32 i;
    int ret = 1;
    fcs = (struct test *) wvMalloc (sizeof (struct test) * nopcd);
    for (i = 0; i < nopcd; i++)
      {
	  fcs[i].fc = pcd[i].fc;
	  fcs[i].offset = (pos[i + 1] - pos[i]) * 2;
      }

    qsort (fcs, nopcd, sizeof (struct test), compar);

    for (i = 0; i < nopcd - 1; i++)
      {
	  if (fcs[i].fc + fcs[i].offset > fcs[i + 1].fc)
	    {
		wvTrace (("overlap, my guess is 8 bit\n"));
		ret = 0;
		break;
	    }
      }

    wvFree (fcs);
    return (ret);
}

```

```c
/* decode_complex.c */


#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <stdlib.h>
#include <stdio.h>
#include "wv.h"

/*
To find the beginning of the paragraph containing a character in a complex
document, it's first necessary to 

1) search for the piece containing the character in the piece table. 

2) Then calculate the FC in the file that stores the character from the piece 
	table information. 
	
3) Using the FC, search the FCs FKP for the largest FC less than the character's 
	FC, call it fcTest. 
	
4) If the character at fcTest-1 is contained in the current piece, then the 
	character corresponding to that FC in the piece is the first character of 
	the paragraph. 
	
5) If that FC is before or marks the beginning of the piece, scan a piece at a 
time towards the beginning of the piece table until a piece is found that 
contains a paragraph mark. 

(This can be done by using the end of the piece FC, finding the largest FC in 
its FKP that is less than or equal to the end of piece FC, and checking to see 
if the character in front of the FKP FC (which must mark a paragraph end) is 
within the piece.)

6) When such an FKP FC is found, the FC marks the first byte of paragraph text.
*/

/*
To find the end of a paragraph for a character in a complex format file,
again 

1) it is necessary to know the piece that contains the character and the
FC assigned to the character. 

2) Using the FC of the character, first search the FKP that describes the 
character to find the smallest FC in the rgfc that is larger than the character 
FC. 

3) If the FC found in the FKP is less than or equal to the limit FC of the 
piece, the end of the paragraph that contains the character is at the FKP FC 
minus 1. 

4) If the FKP FC that was found was greater than the FC of the end of the 
piece, scan piece by piece toward the end of the document until a piece is 
found that contains a paragraph end mark. 

5) It's possible to check if a piece contains a paragraph mark by using the 
FC of the beginning of the piece to search in the FKPs for the smallest FC in 
the FKP rgfc that is greater than the FC of the beginning of the piece. 

If the FC found is less than or equal to the limit FC of the
piece, then the character that ends the paragraph is the character
immediately before the FKP FC.
*/
int
wvGetComplexParaBounds (wvVersion ver, PAPX_FKP * fkp, U32 * fcFirst,
			U32 * fcLim, U32 currentfc, CLX * clx, BTE * bte,
			U32 * pos, int nobte, U32 piece, wvStream * fd)
{
    /*
       U32 currentfc;
     */
    BTE entry;
    long currentpos;

    if (currentfc == 0xffffffffL)
      {
	  wvError (
		   ("Para Bounds not found !, this is ok if this is the last para, otherwise its a disaster\n"));
	  return (-1);
      }

    if (0 != wvGetBTE_FromFC (&entry, currentfc, bte, pos, nobte))
      {
	  wvError (("BTE not found !\n"));
	  return (-1);
      }
    currentpos = wvStream_tell (fd);
    /*The pagenumber of the FKP is entry.pn */

    wvTrace (("the entry.pn is %d\n", entry.pn));
    wvGetPAPX_FKP (ver, fkp, entry.pn, fd);

    wvGetComplexParafcFirst (ver, fcFirst, currentfc, clx, bte, pos, nobte,
			     piece, fkp, fd);

    wvReleasePAPX_FKP (fkp);
    wvTrace (("BREAK\n"));
    wvGetPAPX_FKP (ver, fkp, entry.pn, fd);

    piece =
	wvGetComplexParafcLim (ver, fcLim, currentfc, clx, bte, pos, nobte,
			       piece, fkp, fd);

    wvStream_goto (fd, currentpos);
    return (piece);
}

int
wvGetComplexParafcLim (wvVersion ver, U32 * fcLim, U32 currentfc, CLX * clx,
		       BTE * bte, U32 * pos, int nobte, U32 piece,
		       PAPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest, beginfc;
    BTE entry;
    *fcLim = 0xffffffffL;
    wvTrace (("here is fcLim, currentfc is %x\n", currentfc));
    fcTest = wvSearchNextSmallestFCPAPX_FKP (fkp, currentfc);

    wvTrace (
	     ("fcTest is %x, end is %x\n", fcTest,
	      wvGetEndFCPiece (piece, clx)));


    if (fcTest <= wvGetEndFCPiece (piece, clx))
      {
	  *fcLim = fcTest;
      }
    else
      {
	  /*get end fc of previous piece */
	  piece++;
	  while (piece < clx->nopcd)
	    {
		wvTrace (("piece is %d\n", piece));
		beginfc = wvNormFC (clx->pcd[piece].fc, NULL);
		if (0 != wvGetBTE_FromFC (&entry, beginfc, bte, pos, nobte))
		  {
		      wvError (("BTE not found !\n"));
		      return (-1);
		  }
		wvReleasePAPX_FKP (fkp);
		wvGetPAPX_FKP (ver, fkp, entry.pn, fd);
		fcTest = wvSearchNextSmallestFCPAPX_FKP (fkp, beginfc);
		wvTrace (
			 ("fcTest(t) is %x, end is %x\n", fcTest,
			  wvGetEndFCPiece (piece, clx)));
		if (fcTest <= wvGetEndFCPiece (piece, clx))
		  {
		      *fcLim = fcTest;
		      break;
		  }
		piece++;
	    }
      }
    wvTrace (("fcLim is %x\n", *fcLim));
    if (piece == clx->nopcd)
      {
	  wvTrace (("failed to find a solution to end of paragraph\n"));
	  *fcLim = fcTest;
	  return (clx->nopcd - 1);	/* test using this */
      }
    return (piece);
}


int
wvGetComplexParafcFirst (wvVersion ver, U32 * fcFirst, U32 currentfc,
			 CLX * clx, BTE * bte, U32 * pos, int nobte,
			 U32 piece, PAPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest, endfc;
    BTE entry;
    fcTest = wvSearchNextLargestFCPAPX_FKP (fkp, currentfc);

    wvTrace (("fcTest (s) is %x\n", fcTest));

    if (wvQuerySamePiece (fcTest - 1, clx, piece))
      {
	  wvTrace (("same piece\n"));
	  *fcFirst = fcTest - 1;
      }
    else
      {
	  /*
	     get end fc of previous piece ??, or use the end of the current piece
	   */
	  piece--;
	  while (piece != 0xffffffffL)
	    {
		wvTrace (("piece is %d\n", piece));
		endfc = wvGetEndFCPiece (piece, clx);
		wvTrace (("endfc is %x\n", endfc));
		if (0 != wvGetBTE_FromFC (&entry, endfc, bte, pos, nobte))
		  {
		      wvError (("BTE not found !\n"));
		      return (-1);
		  }
		wvReleasePAPX_FKP (fkp);
		wvGetPAPX_FKP (ver, fkp, entry.pn, fd);
		fcTest = wvSearchNextLargestFCPAPX_FKP (fkp, endfc);
		wvTrace (("fcTest(ft) is %x\n", fcTest));
		if (wvQuerySamePiece (fcTest - 1, clx, piece))
		  {
		      *fcFirst = fcTest - 1;
		      break;
		  }
		piece--;
	    }

      }
    if (piece == 0xffffffffL)
      {
	  wvTrace (
		   ("failed to find a solution to the beginning of the paragraph\n"));
	  *fcFirst = currentfc;
      }
    wvTrace (("fcFirst is finally %x\n", *fcFirst));
    return (0);
}


/* char properties version of the above -JB */
/* only difference is that we're using CHPX FKP pages,
 * and specifically just the Get and Release functions are
 * different between the two. We might be able to 
 * abstract the necessary functions to avoid duplicating them... */

int
wvGetComplexCharBounds (wvVersion ver, CHPX_FKP * fkp, U32 * fcFirst,
			U32 * fcLim, U32 currentfc, CLX * clx, BTE * bte,
			U32 * pos, int nobte, U32 piece, wvStream * fd)
{
    BTE entry;
    long currentpos;

    wvTrace (("current fc is %x\n", currentfc));

    if (currentfc == 0xffffffffL)
      {
	  wvTrace (
		   ("Char Bounds not found !, this is ok if this is the last char, otherwise its a disaster\n"));
	  return (-1);
      }

    if (0 != wvGetBTE_FromFC (&entry, currentfc, bte, pos, nobte))
      {
	  wvError (("BTE not found !\n"));
	  return (-1);
      }
    currentpos = wvStream_tell (fd);
    /*The pagenumber of the FKP is entry.pn */

    wvGetCHPX_FKP (ver, fkp, entry.pn, fd);

    wvGetComplexCharfcFirst (ver, fcFirst, currentfc, clx, bte, pos, nobte,
			     piece, fkp, fd);
    wvTrace (("BEFORE PIECE is %d\n", piece));

    wvReleaseCHPX_FKP (fkp);
    wvGetCHPX_FKP (ver, fkp, entry.pn, fd);

    piece =
	wvGetComplexCharfcLim (ver, fcLim, currentfc, clx, bte, pos, nobte,
			       piece, fkp, fd);
    wvTrace (("AFTER PIECE is %d\n", piece));

    wvStream_goto (fd, currentpos);
    return (piece);
}

int
wvGetComplexCharfcLim (wvVersion ver, U32 * fcLim, U32 currentfc, CLX * clx,
		       BTE * bte, U32 * pos, int nobte, U32 piece,
		       CHPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest;
    /*
       BTE entry;
     */
    *fcLim = 0xffffffffL;
    /* this only works with the initial rgfc array, which is the
     * same for both CHPX and PAPX FKPs */
    fcTest = wvSearchNextSmallestFCPAPX_FKP ((PAPX_FKP *) fkp, currentfc);

    wvTrace (("fcTest is %x\n", fcTest));

    /*
       this single line replaces all the rest, is it conceivable that i overengineered,
       careful rereading of the spec makes no mention of repeating the para process to
       find the boundaries of the exception text runs
     */
    *fcLim = fcTest;
    wvTrace (("fcLim is %x\n", *fcLim));
    if (piece == clx->nopcd)
	return (clx->nopcd - 1);	/* test using this */
    return (piece);
}


int
wvGetComplexCharfcFirst (wvVersion ver, U32 * fcFirst, U32 currentfc,
			 CLX * clx, BTE * bte, U32 * pos, int nobte,
			 U32 piece, CHPX_FKP * fkp, wvStream * fd)
{
    U32 fcTest /*,endfc */ ;
    /*BTE entry; */
    /* this only works with the initial rgfc array, which is the */
    fcTest = wvSearchNextLargestFCCHPX_FKP (fkp, currentfc);

    wvTrace (("fcTest (s) is %x\n", fcTest));

    /*
       this single line replaces all the rest, is it conceivable that i overengineered,
       careful rereading of the spec makes no mention of repeating the para process to
       find the boundaries of the exception text runs
     */
    *fcFirst = fcTest;
    return (0);
}

/*
how this works,
we seek to the beginning of the text, we loop for a count of charaters that is stored in the fib.

the piecetable divides the text up into various sections, we keep track of our location vs
the next entry in that table, when we reach that location, we seek to the position that
the table tells us to go.

there are special cases for coming to the end of a section, and for the beginning and ends of
pages. for the purposes of headers and footers etc.
*/
void
wvDecodeComplex (wvParseStruct * ps)
{
    U32 piececount = 0, i, j, spiece = 0;
    U32 beginfc, endfc;
	U32 stream_size;
    U32 begincp, endcp;
    int ichartype;
    U8  chartype;
    U16 eachchar;
    U32 para_fcFirst, para_fcLim = 0xffffffffL;
    U32 dummy, nextpara_fcLim = 0xffffffffL;
    U32 char_fcFirst, char_fcLim = 0xffffffffL;
    U32 section_fcFirst, section_fcLim = 0xffffffffL;
    U32 comment_cpFirst = 0xffffffffL, comment_cpLim = 0xffffffffL;
    BTE *btePapx = NULL, *bteChpx = NULL;
    U32 *posPapx = NULL, *posChpx = NULL;
    U32 para_intervals, char_intervals, section_intervals, atrd_intervals;
    int cpiece = 0, npiece = 0;
    PAPX_FKP para_fkp;
    PAP apap;
    CHPX_FKP char_fkp;
    CHP achp;
    int para_pendingclose = 0, comment_pendingclose = 0, char_pendingclose =
	0, section_pendingclose = 0;
    int para_dirty = 0, char_dirty = 0, section_dirty = 0;
    SED *sed;
    SEP sep;
    U32 *posSedx;
    ATRD *atrd, *catrd = NULL;
    U32 *posAtrd;
    STTBF grpXstAtnOwners, SttbfAtnbkmk;
    BKF *bkf;
    U32 *posBKF;
    U32 bkf_intervals;
    BKL *bkl;
    U32 *posBKL;
    U32 bkl_intervals;
    wvVersion ver = wvQuerySupported (&ps->fib, NULL);
    external_wvReleasePAPX_FKP ();
    external_wvReleaseCHPX_FKP ();

    /*dop */
    wvGetDOP (ver, &ps->dop, ps->fib.fcDop,
	      ps->fib.lcbDop, ps->tablefd);

#if 0
/* 
this is the versioning name information, the first 22 bytes of each sttbf entry are 
unknown, the rest is a ordinary unicode string, is the time and date and saved by
encoded into the first 22 bytes.
*/
    STTBF versioning;
    if (ver == 0)
      {
	  U16 *str;
	  wvError (("into the versions\n"));
	  wvGetSTTBF (&versioning, ps->fib.fcSttbfUssr, ps->fib.lcbSttbfUssr,
		      ps->tablefd);
	  str = UssrStrBegin (&versioning, 0);
	  wvError (("versioning text is %s\n", wvWideStrToMB (str)));
      }
#endif

    wvGetATRD_PLCF (&atrd, &posAtrd, &atrd_intervals, ps->fib.fcPlcfandRef,
		    ps->fib.lcbPlcfandRef, ps->tablefd);
    wvGetGrpXst (&grpXstAtnOwners, ps->fib.fcGrpXstAtnOwners,
		 ps->fib.lcbGrpXstAtnOwners, ps->tablefd);
    wvTrace (
	     ("offset is %x, len is %d\n", ps->fib.fcSttbfAtnbkmk,
	      ps->fib.lcbSttbfAtnbkmk));
    wvGetSTTBF (&SttbfAtnbkmk, ps->fib.fcSttbfAtnbkmk,
		ps->fib.lcbSttbfAtnbkmk, ps->tablefd);
    wvGetBKF_PLCF (&bkf, &posBKF, &bkf_intervals, ps->fib.fcPlcfAtnbkf,
		   ps->fib.lcbPlcfAtnbkf, ps->tablefd);
    wvGetBKL_PLCF (&bkl, &posBKL, &bkl_intervals, ps->fib.fcPlcfAtnbkl,
           ps->fib.lcbPlcfAtnbkl, ps->fib.fcPlcfAtnbkf, ps->fib.lcbPlcfAtnbkf,
           ps->tablefd);

    /*we will need the stylesheet to do anything useful with layout and look */
    wvGetSTSH (&ps->stsh, ps->fib.fcStshf, ps->fib.lcbStshf, ps->tablefd);

    /* get font list */
    if ((ver == WORD6)
	|| (ver == WORD7))
	wvGetFFN_STTBF6 (&ps->fonts, ps->fib.fcSttbfffn, ps->fib.lcbSttbfffn,
			 ps->tablefd);
    else
	wvGetFFN_STTBF (&ps->fonts, ps->fib.fcSttbfffn, ps->fib.lcbSttbfffn,
			ps->tablefd);

    /*we will need the table of names to answer questions like the name of the doc */
    if ((ver == WORD6)
	|| (ver == WORD7))
      {
	  wvGetSTTBF6 (&ps->anSttbfAssoc, ps->fib.fcSttbfAssoc,
		       ps->fib.lcbSttbfAssoc, ps->tablefd);
	  wvGetSTTBF6 (&ps->Sttbfbkmk, ps->fib.fcSttbfbkmk,
		       ps->fib.lcbSttbfbkmk, ps->tablefd);
      }
    else
      {
	  wvGetSTTBF (&ps->anSttbfAssoc, ps->fib.fcSttbfAssoc,
		      ps->fib.lcbSttbfAssoc, ps->tablefd);
	  wvGetSTTBF (&ps->Sttbfbkmk, ps->fib.fcSttbfbkmk,
		      ps->fib.lcbSttbfbkmk, ps->tablefd);
      }

    /*Extract all the list information that we will need to handle lists later on */
    wvGetLST (&ps->lst, &ps->noofLST, ps->fib.fcPlcfLst, ps->fib.lcbPlcfLst,
	      ps->tablefd);
    wvGetLFO_records (&ps->lfo, &ps->lfolvl, &ps->lvl, &ps->nolfo,
		      &ps->nooflvl, ps->fib.fcPlfLfo, ps->fib.lcbPlfLfo,
		      ps->tablefd);
    /* init the starting list number table */
    if (ps->nolfo)
      {
	  ps->liststartnos = (U32 *) wvMalloc (9 * ps->nolfo * sizeof (U32));
	  ps->listnfcs = (U8 *) wvMalloc (9 * ps->nolfo);
	  ps->finallvl = (LVL *) wvMalloc (9 * ps->nolfo * sizeof (LVL));
	  for (i = 0; i < 9 * ps->nolfo; i++)
	    {
		ps->liststartnos[i] = 0xffffffffL;
		ps->listnfcs[i] = 0xff;
		wvInitLVL (&(ps->finallvl[i]));
	    }
      }
    else
      {
	  ps->liststartnos = NULL;
	  ps->listnfcs = NULL;
	  ps->finallvl = NULL;
      }

    /*Extract Graphic Information */
    wvGetFSPA_PLCF (&ps->fspa, &ps->fspapos, &ps->nooffspa,
		    ps->fib.fcPlcspaMom, ps->fib.lcbPlcspaMom, ps->tablefd);
    wvGetFDOA_PLCF (&ps->fdoa, &ps->fdoapos, &ps->nooffdoa,
		    ps->fib.fcPlcdoaMom, ps->fib.lcbPlcdoaMom, ps->tablefd);

    wvGetCLX (ver, &ps->clx,
	      (U32) ps->fib.fcClx, ps->fib.lcbClx, (U8) ps->fib.fExtChar,
	      ps->tablefd);

    para_fcFirst = char_fcFirst = section_fcFirst =
	wvConvertCPToFC (0, &ps->clx);

#ifdef DEBUG
    if ((ps->fib.ccpFtn) || (ps->fib.ccpHdr))
	wvTrace (("Special ending\n"));
#endif

    /*
       we will need the paragraph and character bounds table to make decisions as 
       to where a table begins and ends
     */
    if ((ver == WORD6)
	|| (ver == WORD7))
      {
	  wvGetBTE_PLCF6 (&btePapx, &posPapx, &para_intervals,
			  ps->fib.fcPlcfbtePapx, ps->fib.lcbPlcfbtePapx,
			  ps->tablefd);
	  wvGetBTE_PLCF6 (&bteChpx, &posChpx, &char_intervals,
			  ps->fib.fcPlcfbteChpx, ps->fib.lcbPlcfbteChpx,
			  ps->tablefd);
      }
    else
      {
	  wvGetBTE_PLCF (&btePapx, &posPapx, &para_intervals,
			 ps->fib.fcPlcfbtePapx, ps->fib.lcbPlcfbtePapx,
			 ps->tablefd);
	  wvGetBTE_PLCF (&bteChpx, &posChpx, &char_intervals,
			 ps->fib.fcPlcfbteChpx, ps->fib.lcbPlcfbteChpx,
			 ps->tablefd);
      }

    wvGetSED_PLCF (&sed, &posSedx, &section_intervals, ps->fib.fcPlcfsed,
		   ps->fib.lcbPlcfsed, ps->tablefd);
    wvTrace (("section_intervals is %d\n", section_intervals));

    wvInitPAPX_FKP (&para_fkp);
    wvInitCHPX_FKP (&char_fkp);

    if(wvHandleDocument (ps, DOCBEGIN))
		goto  finish_processing;

	/*get stream size for bounds checking*/
	stream_size = wvStream_size(ps->mainfd);

    /*for each piece */
    for (piececount = 0; piececount < ps->clx.nopcd; piececount++)
      {
	  ichartype =
	      wvGetPieceBoundsFC (&beginfc, &endfc, &ps->clx, piececount);
	  if(ichartype==-1)
		  break;
	  chartype = (U8) ichartype;
	  /*lvm007@aha.ru fix antiloop: check stream size */
	  if(beginfc>stream_size || endfc>stream_size){
		  wvError (
		   ("Piece Bounds out of range!, its a disaster\n"));
		  continue;
	  }
	  wvStream_goto (ps->mainfd, beginfc);
	  /*lvm007@aha.ru fix antiloop fix*/
	  if(wvGetPieceBoundsCP (&begincp, &endcp, &ps->clx, piececount)==-1)
		  break;
	  wvTrace (
		   ("piece begins at %x and ends just before %x. the char end is %x\n",
		    beginfc, endfc, char_fcLim));

	  /*
	     text that is not in the same piece is not guaranteed to have the same properties as
	     the rest of the exception run, so force a stop and restart of these properties.
	   */
	  char_fcLim = beginfc;

	  for (i = begincp, j = beginfc; (i < endcp /*&& i<ps->fib.ccpText */ );
	       i++, j += wvIncFC (chartype))
	    {
		ps->currentcp = i;
		/* character properties */
		if (j == char_fcLim)
		  {
		      wvHandleElement (ps, CHARPROPEND, (void *) &achp,
				       char_dirty);
		      char_pendingclose = 0;
		  }

		/* comment ending location */
		if (i == comment_cpLim)
		  {
		      wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);
		      comment_pendingclose = 0;
		  }

		/* paragraph properties */
		if (j == para_fcLim)
		  {
		      wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
		      para_pendingclose = 0;
		  }

		/* section properties */
		if (j == section_fcLim)
		  {
		      wvHandleElement (ps, SECTIONEND, (void *) &sep,
				       section_dirty);
		      section_pendingclose = 0;
		  }

		if ((section_fcLim == 0xffffffff) || (section_fcLim == j))
		  {
		      section_dirty =
			  wvGetSimpleSectionBounds (ver, ps,
						    &sep, &section_fcFirst,
						    &section_fcLim, i,
						    &ps->clx, sed, &spiece,
						    posSedx,
						    section_intervals,
						    &ps->stsh, ps->mainfd);
		      section_dirty =
			  (wvGetComplexSEP
			   (ver, &sep, spiece,
			    &ps->stsh, &ps->clx) ? 1 : section_dirty);
		  }

		if (j == section_fcFirst)
		  {
		      wvHandleElement (ps, SECTIONBEGIN, (void *) &sep,
				       section_dirty);
		      section_pendingclose = 1;
		  }


		if ((para_fcLim == 0xffffffffL) || (para_fcLim == j))
		  {
		      wvReleasePAPX_FKP (&para_fkp);
		      wvTrace (
			       ("cp and fc are %x(%d) %x\n", i, i,
				wvConvertCPToFC (i, &ps->clx)));
		      cpiece =
			  wvGetComplexParaBounds (ver, &para_fkp,
						  &para_fcFirst, &para_fcLim,
						  wvConvertCPToFC (i,
								   &ps->clx),
						  &ps->clx, btePapx, posPapx,
						  para_intervals, piececount,
						  ps->mainfd);
		      wvTrace (
			       ("para begin and end is %x %x\n", para_fcFirst,
				para_fcLim));

		      if (0 == para_pendingclose)
			{
			    /*
			       if there's no paragraph open, but there should be then I believe that the fcFirst search
			       has failed me, so I set it to now. I need to investigate this further. I believe it occurs
			       when a the last piece ended simultaneously with the last paragraph, and that the algorithm
			       for finding the beginning of a para breaks under that condition. I need more examples to
			       be sure, but it happens is very large complex files so its hard to find
			     */
			    if (j != para_fcFirst)
			      {
				  wvWarning (
					     ("There is no paragraph due to open but one should be, plugging the gap.\n"));
				  para_fcFirst = j;
			      }
			}

		  }

		if (j == para_fcFirst)
		  {
		      para_dirty =
			  wvAssembleSimplePAP (ver, &apap, para_fcLim, &para_fkp, ps);
		      para_dirty =
			  (wvAssembleComplexPAP
			   (ver, &apap, cpiece, ps) ? 1 : para_dirty);
#ifdef SPRMTEST
		      {
			  int p;
			  wvTrace (("Assembled Complex\n"));
			  for (p = 0; p < apap.itbdMac; p++)
			      wvError (
				       ("Tab stop positions are %f inches (%d)\n",
					((float) (apap.rgdxaTab[p])) / 1440,
					apap.rgdxaTab[p]));
		      }
#endif

		      /* test section */
		      wvReleasePAPX_FKP (&para_fkp);
		      wvTrace (
			       ("cp and fc are %x(%d) %x\n", i, i,
				wvConvertCPToFC (i, &ps->clx)));
		      npiece =
			  wvGetComplexParaBounds (ver, &para_fkp,
						  &dummy, &nextpara_fcLim,
						  para_fcLim, &ps->clx,
						  btePapx, posPapx,
						  para_intervals, piececount,
						  ps->mainfd);
		      wvTrace (
			       ("para begin and end is %x %x\n", para_fcFirst,
				para_fcLim));
		      if (npiece > -1)
			{
			    wvAssembleSimplePAP (ver, &ps->nextpap, nextpara_fcLim, &para_fkp, ps);
			    wvAssembleComplexPAP (ver, &ps->nextpap, npiece,ps);
			}
		      else
			  wvInitPAP (&ps->nextpap);
		      /* end test section */

		      if ((apap.fInTable) && (!apap.fTtp))
			{
			    wvGetComplexFullTableInit (ps, para_intervals,
						       btePapx, posPapx,
						       piececount);
			    wvGetComplexRowTap (ps, &apap, para_intervals,
						btePapx, posPapx, piececount);
			}
		      else if (apap.fInTable == 0)
			  ps->intable = 0;

		      wvHandleElement (ps, PARABEGIN, (void *) &apap,
				       para_dirty);

		      char_fcLim = j;
		      para_pendingclose = 1;
		  }


		if ((comment_cpLim == 0xffffffffL) || (comment_cpLim == i))
		  {
		      wvTrace (
			       ("searching for the next comment begin cp is %d\n",
				i));
		      catrd =
			  wvGetCommentBounds (&comment_cpFirst,
					      &comment_cpLim, i, atrd,
					      posAtrd, atrd_intervals,
					      &SttbfAtnbkmk, bkf, posBKF,
					      bkf_intervals, bkl, posBKL,
					      bkl_intervals);
		      wvTrace (
			       ("begin and end are %d %d\n", comment_cpFirst,
				comment_cpLim));
		  }

		if (i == comment_cpFirst)
		  {
		      wvHandleElement (ps, COMMENTBEGIN, (void *) catrd, 0);
		      comment_pendingclose = 1;
		  }


		if ((char_fcLim == 0xffffffffL) || (char_fcLim == j))
		  {
		      wvReleaseCHPX_FKP (&char_fkp);
		      /*try this without using the piece of the end char for anything */
		      wvGetComplexCharBounds (ver, &char_fkp,
					      &char_fcFirst, &char_fcLim,
					      wvConvertCPToFC (i, &ps->clx),
					      &ps->clx, bteChpx, posChpx,
					      char_intervals, piececount,
					      ps->mainfd);
		      wvTrace (
			       ("Bounds from %x to %x\n", char_fcFirst,
				char_fcLim));
		      if (char_fcLim == char_fcFirst)
			  wvError (
				   ("I believe that this is an error, and you might see incorrect character properties\n"));
		      if (0 == char_pendingclose)
			{
			    /*
			       if there's no character run open, but there should be then I believe that the fcFirst search
			       has failed me, so I set it to now. I need to investigate this further.
			     */
			    if (j != char_fcFirst)
			      {
				  wvTrace (
					   ("There is no character run due to open but one should be, plugging the gap.\n"));
				  char_fcFirst = j;
			      }

			}
		      else{
  			 /* lvm007@aha.ru fix: if currentfc>fcFirst but CHARPROP's changed look examples/charprops.doc for decode_simple*/
			 if(char_fcFirst< j)
				char_fcFirst = j;
		       }
		  }

		if (j == char_fcFirst)
		  {
		      /* a CHP's base style is in the para style */
		      /*achp.istd = apap.istd;*/
		      wvTrace (("getting chp\n"));
		      char_dirty =
				  wvAssembleSimpleCHP (ver, &achp, &apap,
					       char_fcLim, &char_fkp,
					       &ps->stsh);
		      wvTrace (("getting complex chp\n"));
		      char_dirty =
			  (wvAssembleComplexCHP
			   (ver, &achp, cpiece,
			    &ps->stsh, &ps->clx) ? 1 : char_dirty);
		      wvHandleElement (ps, CHARPROPBEGIN, (void *) &achp,
				       char_dirty);
		      char_pendingclose = 1;
		  }


		eachchar = wvGetChar (ps->mainfd, chartype);

		/* previously, in place of ps there was a NULL,
		 * but it was crashing Abiword. Was it NULL for a
		 * reason? -JB */
		/* 
		   nah, it was a oversight from when i didn't actually
		   use ps in this function
		   C.
		 */
		if ((eachchar == 0x07) && (!achp.fSpec))
		    ps->endcell = 1;

		wvTrace (("char pos is %x %x\n", j, eachchar));
		wvOutputTextChar (eachchar, chartype, ps, &achp);
	    }

	  if (j == para_fcLim)
	    {
		wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
		para_pendingclose = 0;
		para_fcLim = 0xffffffffL;
	    }

	  if (i == comment_cpLim)
	    {
		wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);
		comment_pendingclose = 0;
		comment_cpLim = 0xffffffffL;
	    }

	  if (j == char_fcLim)
	    {
		wvHandleElement (ps, CHARPROPEND, (void *) &achp, char_dirty);
		char_pendingclose = 0;
		char_fcLim = 0xffffffffL;
	    }

#if 0
	  /*      
	     I might have to rethink this closing tag enforcer for complex mode, have to think the
	     flow out a bit more, this section one is plain wrong, im leaving it here so i won't
	     forget and be tempted to put it back in :-)
	     if (j == section_fcLim)
	     {
	     wvHandleElement(ps, SECTIONEND, (void*)&sep,section_dirty);
	     section_pendingclose=0;
	     }
	   */
#endif
      }

 finish_processing:
    if (char_pendingclose)
      {
	  wvInitCHP (&achp);
	  wvHandleElement (ps, CHARPROPEND, (void *) &achp, char_dirty);
      }

    if (comment_pendingclose)
	wvHandleElement (ps, COMMENTEND, (void *) catrd, 0);

    if (para_pendingclose)
      {
	  wvInitPAP (&apap);
	  wvHandleElement (ps, PARAEND, (void *) &apap, para_dirty);
      }

    if (section_pendingclose)
	wvHandleElement (ps, SECTIONEND, (void *) &sep, section_dirty);

    wvFree (ps->fspa);
    wvFree (ps->fspapos);
    wvFree (ps->fdoa);
    wvFree (ps->fdoapos);

    wvFree (posBKL);
    wvFree (bkl);
    wvFree (posBKF);
    wvFree (bkf);
    wvFree (posAtrd);
    wvFree (atrd);

    wvReleasePAPX_FKP (&para_fkp);
    wvReleaseCHPX_FKP (&char_fkp);

    wvHandleDocument (ps, DOCEND);
    wvFree (posSedx);
    wvFree (sed);

    wvFree (ps->liststartnos);
    wvFree (ps->listnfcs);
    for (i = 0; i < 9 * ps->nolfo; i++)
	wvReleaseLVL (&(ps->finallvl[i]));
    wvFree (ps->finallvl);

    wvReleaseLST (&ps->lst, ps->noofLST);
    wvReleaseLFO_records (&ps->lfo, &ps->lfolvl, &ps->lvl, ps->nooflvl);
    wvReleaseSTTBF (&ps->anSttbfAssoc);

    wvFree (btePapx);
    wvFree (posPapx);
    wvFree (bteChpx);
    wvFree (posChpx);
    wvReleaseCLX (&ps->clx);
    wvReleaseFFN_STTBF (&ps->fonts);
    wvReleaseSTSH (&ps->stsh);
    wvReleaseSTTBF (&SttbfAtnbkmk);
    wvReleaseSTTBF (&grpXstAtnOwners);
    if (ps->vmerges)
      {
	  for (i = 0; i < ps->norows; i++)
	      wvFree (ps->vmerges[i]);
	  wvFree (ps->vmerges);
      }
    wvFree (ps->cellbounds);
	wvOLEFree(ps);
    tokenTreeFreeAll ();
}

/*
 The process thus far has created a SEP that describes what the section properties of 
 the section at the last full save. 

 1) Now apply any section sprms that were linked to the piece that contains the 
 section's section mark. 
 
 2) If pcd.prm.fComplex is 0, pcd.prm contains 1 sprm which should be applied to 
 the local SEP if it is a section sprm. 
 
 3) If pcd.prm.fComplex is 1, pcd.prm.igrpprl is the index of a grpprl in the CLX. 
 If that grpprl contains any section sprms, they should be applied to the local SEP
*/
int
wvGetComplexSEP (wvVersion ver, SEP * sep, U32 cpiece, STSH * stsh, CLX * clx)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (clx->pcd[cpiece].prm.fComplex == 0)
      {
	  val = clx->pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton\n", clx->pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) clx->pcd[cpiece].prm.
						     para.var1.isprm), NULL,
				     NULL, sep, stsh, pointer, &pos, NULL);
	  if (RetSprm.sgc == sgcSep)
	      ret = 1;
      }
    else
      {
	  index = clx->pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  fprintf (stderr, "\n");
	  while (i < clx->cbGrpprl[index])
	    {
		fprintf (stderr, "%x (%d)\n", *(clx->grpprl[index] + i),
			 *(clx->grpprl[index] + i));
		i++;
	    }
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < clx->cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (clx->grpprl[index] + i, &i);
		else
		  {
		      sprm = bread_8ubit (clx->grpprl[index] + i, &i);
		      sprm = (U8) wvGetrgsprmWord6 ( (U8) sprm);
		  }
		pointer = clx->grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, NULL, NULL, sep, stsh,
					   pointer, &i, NULL);
		if (RetSprm.sgc == sgcSep)
		    ret = 1;
	    }
      }
    return (ret);
}

/*
The process thus far has created a PAP that describes
what the paragraph properties of the paragraph were at the last full save.

1) Now it's necessary to apply any paragraph sprms that were linked to the
piece that contains the paragraph's paragraph mark. 

2) If pcd.prm.fComplex is 0, pcd.prm contains 1 sprm which should only be 
applied to the local PAP if it is a paragraph sprm. 

3) If pcd.prm.fComplex is 1, pcd.prm.igrpprl is the index of a grpprl in the 
CLX.  If that grpprl contains any paragraph sprms, they should be applied to 
the local PAP.
*/
int
wvAssembleComplexPAP (wvVersion ver, PAP * apap, U32 cpiece, wvParseStruct *ps)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 sprm8;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (ps->clx.pcd[cpiece].prm.fComplex == 0)
      {
	  val = ps->clx.pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton\n", ps->clx.pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) ps->clx.pcd[cpiece].prm.
						     para.var1.isprm), apap,
				     NULL, NULL, &ps->stsh, pointer, &pos, ps->data);
	  if (RetSprm.sgc == sgcPara)
	      ret = 1;
      }
    else
      {
	  index = ps->clx.pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  wvError (("HERE-->\n"));
	  fprintf (stderr, "\n");
	  for (i = 0; i < ps->clx.cbGrpprl[index]; i++)
	      fprintf (stderr, "%x ", *(ps->clx.grpprl[index] + i));
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < ps->clx.cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (ps->clx.grpprl[index] + i, &i);
		else
		  {
		      sprm8 = bread_8ubit (ps->clx.grpprl[index] + i, &i);
		      sprm = (U16) wvGetrgsprmWord6 (sprm8);
		      wvTrace (("sprm is %x\n", sprm));
		  }
		pointer = ps->clx.grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, apap, NULL, NULL, &ps->stsh,
					   pointer, &i, ps->data);
		if (RetSprm.sgc == sgcPara)
		    ret = 1;
	    }
      }
    return (ret);
}

/* CHP version of the above. follows the same rules -JB */
int
wvAssembleComplexCHP (wvVersion ver, CHP * achp, U32 cpiece, STSH * stsh,
		      CLX * clx)
{
    int ret = 0;
    U16 sprm, pos = 0, i = 0;
    U8 sprm8;
    U8 *pointer;
    U16 index;
    U8 val;
    Sprm RetSprm;

    if (clx->pcd[cpiece].prm.fComplex == 0)
      {
	  val = clx->pcd[cpiece].prm.para.var1.val;
	  pointer = &val;
#ifdef SPRMTEST
	  wvError (("singleton %d\n", clx->pcd[cpiece].prm.para.var1.isprm));
#endif
	  RetSprm =
	      wvApplySprmFromBucket (ver,
				     (U16) wvGetrgsprmPrm ( (U16) clx->pcd[cpiece].prm.
						     para.var1.isprm), NULL,
				     achp, NULL, stsh, pointer, &pos, NULL);
	  if (RetSprm.sgc == sgcChp)
	      ret = 1;
      }
    else
      {
	  index = clx->pcd[cpiece].prm.para.var2.igrpprl;
#ifdef SPRMTEST
	  fprintf (stderr, "\n");
	  for (i = 0; i < clx->cbGrpprl[index]; i++)
	      fprintf (stderr, "%x ", *(clx->grpprl[index] + i));
	  fprintf (stderr, "\n");
	  i = 0;
#endif
	  while (i < clx->cbGrpprl[index])
	    {
		if (ver == WORD8)
		    sprm = bread_16ubit (clx->grpprl[index] + i, &i);
		else
		  {
		      sprm8 = bread_8ubit (clx->grpprl[index] + i, &i);
		      sprm = (U16) wvGetrgsprmWord6 (sprm8);
		  }
		pointer = clx->grpprl[index] + i;
		RetSprm =
		    wvApplySprmFromBucket (ver, sprm, NULL, achp, NULL, stsh,
					   pointer, &i, NULL);
		if (RetSprm.sgc == sgcChp)
		    ret = 1;
	    }
      }
    return (ret);
}

```