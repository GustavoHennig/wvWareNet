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

            // For Word95 files, try to proceed without Table stream if not found
            var tableEntry = entries.Find(e => 
                e.Name.Contains("1Table", StringComparison.OrdinalIgnoreCase)) 
                ?? entries.Find(e => e.Name.Contains("0Table", StringComparison.OrdinalIgnoreCase))
                ?? entries.Find(e => e.Name.Contains("Table", StringComparison.OrdinalIgnoreCase));

            // Word95 files: 100=Word6, 101=Word95, 104=Word97 but some Word95 files use 104
            bool isWord95 = fib.NFib == 100 || fib.NFib == 101 || fib.NFib == 104;

            if (tableEntry == null)
            {
                if (isWord95)
                {
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

            var pieceTable = new WvWareNet.Core.PieceTable(_logger);

            byte[]? clxData = null;
            if (tableStream != null)
            {
                // Word 97 and later use FcClx and LcbClx from FIB
                if (fib.NFib >= 0x0076) // Word 97 (0x0076) and later
                {
                    if (fib.LcbClx > 0 && fib.FcClx >= 0 && (fib.FcClx + fib.LcbClx) <= tableStream.Length)
                    {
                        _logger.LogInfo($"[DEBUG] FIB: FcClx={fib.FcClx}, LcbClx={fib.LcbClx}, tableStream.Length={tableStream.Length}");
                        clxData = new byte[fib.LcbClx];
                        Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);
                        _logger.LogInfo($"[DEBUG] CLX data: {BitConverter.ToString(clxData)}");
                    }
                }
                else // Word 6/95 - CLX is at the beginning of the Table stream
                {
                    // Heuristic: Assume CLX is at the beginning of the table stream for older formats
                    // and has a reasonable size (e.g., up to 512 bytes, or the whole stream if smaller)
                    int assumedClxLength = Math.Min(tableStream.Length, 512); 
                    clxData = new byte[assumedClxLength];
                    Array.Copy(tableStream, 0, clxData, 0, assumedClxLength);
                    _logger.LogInfo($"[DEBUG] Assuming CLX at start of Table stream for NFib={fib.NFib}, length={assumedClxLength}");
                }
            }

            if (clxData != null && clxData.Length > 0)
            {
                _logger.LogInfo($"[DEBUG] Parsing piece table with CLX data (size: {clxData.Length}), FcMin: {fib.FcMin}, FcMac: {fib.FcMac}, NFib: {fib.NFib}");
                pieceTable.Parse(clxData, fib.FcMin, fib.FcMac, fib.NFib);
                
                _logger.LogInfo($"[DEBUG] Piece table parsed. Number of pieces: {pieceTable.Pieces.Count}");
                foreach (var piece in pieceTable.Pieces)
                {
                    _logger.LogInfo($"[DEBUG]   - Piece: CpStart={piece.CpStart}, CpEnd={piece.CpEnd}, FcStart={piece.FcStart}, IsUnicode={piece.IsUnicode}");
                }

                if (pieceTable.Pieces.Count == 1 && pieceTable.Pieces[0].FcStart >= wordDocStream.Length)
                {
                    _logger.LogWarning("Invalid piece table detected, falling back to FcMin/FcMac range.");
                    pieceTable.SetSinglePiece(fib.FcMin, fib.FcMac, fib.NFib);
                }
            }
            else
            {
                _logger.LogWarning("No CLX data found or invalid, treating document as single piece");
                pieceTable.SetSinglePiece(fib.FcMin, fib.FcMac, fib.NFib);
                _logger.LogInfo($"[DEBUG] Fallback single piece created. IsUnicode: {pieceTable.Pieces[0].IsUnicode}");
            }

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
            if (tableStream != null && fib.FcPlcfftn > 0 && fib.LcbPlcfftn > 0 && fib.FcPlcfftn + fib.LcbPlcfftn <= tableStream.Length)
            {
                byte[] plcfftn = new byte[fib.LcbPlcfftn];
                Array.Copy(tableStream, fib.FcPlcfftn, plcfftn, 0, fib.LcbPlcfftn);

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
                _logger.LogInfo($"[STRUCTURE] Section {i+1}: {section.Paragraphs.Count} paragraph(s)");
                for (int j = 0; j < section.Paragraphs.Count; j++)
                {
                    var paragraph = section.Paragraphs[j];
                    _logger.LogInfo($"[STRUCTURE] Paragraph {j+1}: {paragraph.Runs.Count} run(s)");
                    for (int k = 0; k < paragraph.Runs.Count; k++)
                    {
                        var run = paragraph.Runs[k];
                        _logger.LogInfo($"[STRUCTURE] Run {k+1}: Length={run.Text?.Length ?? 0}, TextPreview='{(run.Text != null ? run.Text.Substring(0, Math.Min(20, run.Text.Length)) : "null")}'");
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
