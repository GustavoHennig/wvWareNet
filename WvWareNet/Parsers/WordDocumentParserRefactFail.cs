using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WvWareNet.Parsers;
using WvWareNet.Utilities;

namespace WvWareNet.Parsers
{

    /// <summary>
    /// Just for reference, this is a failed refactor attempt.
    /// </summary>
    public class WordDocumentParserRefactFail
    {
        private readonly CompoundFileBinaryFormatParser _cfbfParser;
        private readonly ILogger _logger;
        private Core.DocumentModel _documentModel;

        public WordDocumentParserRefactFail(CompoundFileBinaryFormatParser cfbfParser, ILogger logger)
        {
            _cfbfParser = cfbfParser;
            _logger = logger;
            _documentModel = new Core.DocumentModel();
        }

        private static bool IsAlphaNumeric(byte b)
        {
            char c = (char)b;
            return char.IsLetterOrDigit(c);
        }

        // Implements [MS-DOC] 2.4.2 paragraph boundary algorithm
        // Returns the character position of the first character in the paragraph containing cp
        private int FindParagraphStartCp(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream)
        {
            if (cp >= fib.CcpText)
                return (int)fib.CcpText - 1;

            // If no PlcBtePapx data available, fall back to beginning of document
            if (fib.FcPlcfbtePapx <= 0 || fib.LcbPlcfbtePapx <= 0 || tableStream == null)
                return 0;

            // Prevent infinite loops with recursion depth limit
            int maxDepth = 50;
            return FindParagraphStartCpInternal(cp, fib, pieceTable, tableStream, wordDocStream, 0, maxDepth);
        }

        private int FindParagraphStartCpInternal(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream, int depth, int maxDepth)
        {
            if (depth >= maxDepth)
            {
                _logger.LogWarning($"Maximum recursion depth reached in FindParagraphStartCp. Returning 0.");
                return 0;
            }

            int pieceIndex = pieceTable.GetPieceIndexFromCp(cp);
            if (pieceIndex == -1)
                return 0;

            var piece = pieceTable.Pieces[pieceIndex];

            // Convert CP to FC using the new method
            int fc = pieceTable.ConvertCpToFc(cp);
            if (fc == -1)
                return 0;

            // Step 4: Read PlcBtePapx from Table Stream
            byte[] plcbtePapx = new byte[fib.LcbPlcfbtePapx];
            Array.Copy(tableStream, fib.FcPlcfbtePapx, plcbtePapx, 0, fib.LcbPlcfbtePapx);

            // PlcBtePapx: aFc[] (int32), aPnBtePapx[] (uint16), last aFc is end
            int nPapx = (int)((fib.LcbPlcfbtePapx - 4) / 6);
            if (nPapx <= 0) return 0;

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
                if (fcLast < piece.FcStart)
                {
                    // Step 8: If PlcPcd.aCp[i] is 0, return 0
                    if (piece.CpStart == 0)
                        return 0;
                    // Step 9: Set cp = PlcPcd.aCp[i], i = i-1, repeat
                    return FindParagraphStartCpInternal(piece.CpStart, fib, pieceTable, tableStream, wordDocStream, depth + 1, maxDepth);
                }
                fc = fcLast;
                // Step 7: If fcFirst > fcPcd, compute dfc and return paragraph start cp
                int dfc = fcLast - piece.FcStart;
                if (piece.IsUnicode)
                    dfc /= 2;
                int paraStartCp = piece.CpStart + dfc;
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
            if (papxFkpOffset >= wordDocStream.Length)
                return 0;

            wordDocStream.Seek(papxFkpOffset, System.IO.SeekOrigin.Begin);
            byte[] papxFkp = new byte[512];
            int read = wordDocStream.Read(papxFkp, 0, Math.Min(512, (int)(wordDocStream.Length - papxFkpOffset)));

            // If less than 512 bytes, pad the rest with zeros
            if (read < 512)
            {
                for (int z = read; z < 512; z++)
                    papxFkp[z] = 0;
            }

            // PapxFkp: first n+1 int32 rgfc, then n Papx, then crun byte at 511
            int crun = papxFkp[511];
            if (read == 0 || crun == 0)
                return 0;

            int[] rgfc = new int[crun + 1];
            for (int k = 0; k < crun + 1; k++)
                rgfc[k] = BitConverter.ToInt32(papxFkp, k * 4);

            // Handle empty or invalid PapxFkp block
            if (rgfc.Length == 0 || rgfc[^1] == 0)
                return 0;

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
                return (int)fib.CcpText - 1;

            int fcFirst = rgfc[kIdx];

            // Step 7: If fcFirst > fcPcd, compute dfc and return paragraph start cp
            if (fcFirst > piece.FcStart)
            {
                int dfc = fcFirst - piece.FcStart;
                if (piece.IsUnicode)
                    dfc /= 2;
                int paraStartCp = piece.CpStart + dfc;
                return paraStartCp;
            }
            // Step 8: If PlcPcd.aCp[i] is 0, return 0
            if (piece.CpStart == 0)
                return 0;
            // Step 9: Set cp = PlcPcd.aCp[i], i = i-1, repeat
            return FindParagraphStartCpInternal(piece.CpStart, fib, pieceTable, tableStream, wordDocStream, depth + 1, maxDepth);
        }

        // Implements [MS-DOC] 2.4.2 paragraph boundary algorithm
        // Returns the character position of the last character in the paragraph containing cp
        private int FindParagraphEndCp(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream)
        {
            // If no PlcBtePapx data available, return end of text
            if (fib.FcPlcfbtePapx <= 0 || fib.LcbPlcfbtePapx <= 0 || tableStream == null)
                return (int)fib.CcpText - 1;

            // Prevent infinite loops with recursion depth limit
            int maxDepth = 50;
            return FindParagraphEndCpInternal(cp, fib, pieceTable, tableStream, wordDocStream, 0, maxDepth);
        }

        private int FindParagraphEndCpInternal(int cp, WvWareNet.Core.FileInformationBlock fib, WvWareNet.Core.PieceTable pieceTable, byte[] tableStream, System.IO.Stream wordDocStream, int depth, int maxDepth)
        {
            if (depth >= maxDepth)
            {
                _logger.LogWarning($"Maximum recursion depth reached in FindParagraphEndCp. Returning end of text.");
                return (int)fib.CcpText - 1;
            }

            int pieceIndex = pieceTable.GetPieceIndexFromCp(cp);
            if (pieceIndex == -1)
                return (int)fib.CcpText - 1;

            var piece = pieceTable.Pieces[pieceIndex];

            // Convert CP to FC using the new method
            int fc = pieceTable.ConvertCpToFc(cp);
            if (fc == -1)
                return (int)fib.CcpText - 1;

            int fcMac = pieceTable.ConvertCpToFc(piece.CpEnd - 1);
            if (fcMac == -1)
                fcMac = piece.FcStart + (piece.IsUnicode ? (piece.CpEnd - piece.CpStart) * 2 : (piece.CpEnd - piece.CpStart));

            // Step 4: Read PlcBtePapx from Table Stream
            byte[] plcbtePapx = new byte[fib.LcbPlcfbtePapx];
            Array.Copy(tableStream, fib.FcPlcfbtePapx, plcbtePapx, 0, fib.LcbPlcfbtePapx);

            int nPapx = (int)((fib.LcbPlcfbtePapx - 4) / 6);
            if (nPapx <= 0) return (int)fib.CcpText - 1;

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
            if (papxFkpOffset >= wordDocStream.Length)
                return (int)fib.CcpText - 1;

            wordDocStream.Seek(papxFkpOffset, System.IO.SeekOrigin.Begin);
            byte[] papxFkp = new byte[512];
            int read = wordDocStream.Read(papxFkp, 0, Math.Min(512, (int)(wordDocStream.Length - papxFkpOffset)));

            if (read < 512)
            {
                for (int z = read; z < 512; z++)
                    papxFkp[z] = 0;
            }

            int crun = papxFkp[511];
            if (read == 0 || crun == 0)
                return (int)fib.CcpText - 1;

            int[] rgfc = new int[crun + 1];
            for (int k = 0; k < crun + 1; k++)
                rgfc[k] = BitConverter.ToInt32(papxFkp, k * 4);

            if (rgfc.Length == 0 || rgfc[^1] == 0)
                return (int)fib.CcpText - 1;

            // Step 5: Find largest k such that rgfc[k] <= fc
            int kIdx = 0;
            for (int k = 0; k < rgfc.Length; k++)
            {
                if (rgfc[k] <= fc)
                    kIdx = k;
                else
                    break;
            }
            if (rgfc[^1] <= fc || kIdx >= rgfc.Length - 1)
                return (int)fib.CcpText - 1;

            int fcLim = rgfc[kIdx + 1];

            // Step 6: If fcLim <= fcMac, compute dfc and return paragraph end cp
            if (fcLim <= fcMac)
            {
                int dfc = fcLim - piece.FcStart;
                if (piece.IsUnicode)
                    dfc /= 2;
                int paraEndCp = piece.CpStart + dfc - 1;
                return paraEndCp;
            }
            // Step 7: Set cp = PlcPcd.aCp[i+1], i = i+1, repeat
            int nextCp = piece.CpEnd;
            if (pieceIndex + 1 >= pieceTable.Pieces.Count)
                return pieceTable.Pieces[^1].CpEnd - 1;
            return FindParagraphEndCpInternal(nextCp, fib, pieceTable, tableStream, wordDocStream, depth + 1, maxDepth);
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

            // Read stream data
            var wordDocStream = _cfbfParser.ReadStream(wordDocEntry);

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

            var streamName = fib.FWhichTblStm ? "1Table" : "0Table";
            var tableEntry = entries.FirstOrDefault(e =>
                e.Name.Equals(streamName, StringComparison.OrdinalIgnoreCase));

            // Word95 files: 100=Word6, 101=Word95, 104=Word97 but some Word95 files use 104
            bool isWord95 = fib.NFib == 100 || fib.NFib == 101 || fib.NFib == 104;

            byte[]? tableStream = null;
            if (tableEntry != null)
            {
                _logger.LogInfo($"[DEBUG] Found Table stream: {tableEntry.Name}");
                tableStream = _cfbfParser.ReadStream(tableEntry);
            }
            else
            {
                // For documents without table stream, we can still proceed with limited functionality
                _logger.LogWarning($"Table stream not found (NFib={fib.NFib}), attempting to parse with reduced functionality");
            }

            _documentModel.FileInfo = fib;

            if ((fib.FEncrypted || fib.FCrypto))
                throw new NotSupportedException("Encrypted Word documents are not supported.");

            if (fib.FibVersion == null)
                _logger.LogWarning($"Unknown Word version NFib={fib.NFib}");
            else
                _logger.LogInfo($"Detected Word version: {fib.FibVersion}");

            var pieceTable = WvWareNet.Core.PieceTable.CreateFromStreams(_logger, fib, tableStream, wordDocStream);

            // Parse stylesheet
            WvWareNet.Core.Stylesheet stylesheet = new WvWareNet.Core.Stylesheet();
            if (tableStream != null && fib.FcStshf > 0 && fib.LcbStshf > 0 && fib.FcStshf + fib.LcbStshf <= tableStream.Length)
            {
                byte[] stshData = new byte[fib.LcbStshf];
                Array.Copy(tableStream, fib.FcStshf, stshData, 0, fib.LcbStshf);
                stylesheet = WvWareNet.Core.Stylesheet.Parse(stshData);
                _logger.LogInfo($"[DEBUG] Parsed stylesheet with {stylesheet.Styles.Count} styles");
            }
            else
            {
                _logger.LogWarning($"No stylesheet found - using default built-in styles");
                // Create basic default stylesheet
                stylesheet = new WvWareNet.Core.Stylesheet();
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 0, Name = "Normal" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 1, Name = "heading 1" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 2, Name = "heading 2" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 3, Name = "heading 3" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 4, Name = "heading 4" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 5, Name = "heading 5" });
                stylesheet.Styles.Add(new WvWareNet.Core.Style { Index = 6, Name = "heading 6" });
            }

            // Extract text using PieceTable and populate DocumentModel with sections and paragraphs
            _documentModel = new WvWareNet.Core.DocumentModel();
            _documentModel.FileInfo = fib;
            _documentModel.Stylesheet = stylesheet;

            var defaultSection = new WvWareNet.Core.Section();
            _documentModel.Sections.Add(defaultSection);

            using var wordDocMs = new System.IO.MemoryStream(wordDocStream);

            // Try to use paragraph boundary detection if possible
            bool useParagraphBoundaries = tableStream != null &&
                                            fib.FcPlcfbtePapx > 0 &&
                                            fib.LcbPlcfbtePapx > 0 &&
                                            fib.FcPlcfbtePapx + fib.LcbPlcfbtePapx <= tableStream.Length;

            if (pieceTable.Pieces.Count <= 1 || (fib.LcbPlcfbtePapx > 0 && fib.LcbPlcfbtePapx <= 16))
            {
                _logger.LogWarning("Suspicious piece table or PLCBTEPapx size, forcing fallback to simple extraction.");
                useParagraphBoundaries = false;
            }

            if (useParagraphBoundaries)
            {
                try
                {
                    int cp = 0;
                    int paraIdx = 0;
                    int maxParagraphs = 10000; // Safety limit
                    int consecutiveErrors = 0;

                    while (cp < fib.CcpText && paraIdx < maxParagraphs && consecutiveErrors < 10)
                    {
                        int paraStart = FindParagraphStartCp(cp, fib, pieceTable, tableStream, wordDocMs);
                        int paraEnd = FindParagraphEndCp(cp, fib, pieceTable, tableStream, wordDocMs);

                        // Validate boundaries
                        if (paraEnd < paraStart || paraStart < 0 || paraEnd < 0)
                        {
                            _logger.LogWarning($"Invalid paragraph boundaries at CP={cp}. Start={paraStart}, End={paraEnd}");
                            consecutiveErrors++;
                            cp++;
                            continue;
                        }

                        consecutiveErrors = 0; // Reset error counter on success

                        if (paraEnd >= fib.CcpText)
                            paraEnd = (int)fib.CcpText - 1;

                        var paragraph = new WvWareNet.Core.Paragraph();
                        paragraph.StyleIndex = 0;
                        paragraph.Style = stylesheet.GetStyleName(0);

                        _logger.LogInfo($"[DEBUG] Paragraph {paraIdx}: CP {paraStart}-{paraEnd}");

                        // Extract text for this paragraph
                        try
                        {
                            string text = pieceTable.GetTextForRange(paraStart, paraEnd + 1, wordDocMs);
                            if (!string.IsNullOrEmpty(text))
                            {
                                var run = new WvWareNet.Core.Run { Text = text };
                                paragraph.Runs.Add(run);
                                defaultSection.Paragraphs.Add(paragraph);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning($"Failed to extract text for paragraph {paraIdx}: {ex.Message}");
                        }

                        paraIdx++;

                        // Ensure we make progress
                        int nextCp = paraEnd + 1;
                        if (nextCp <= cp)
                        {
                            _logger.LogWarning($"No progress made at CP={cp}. Forcing fallback to simple extraction.");
                            useParagraphBoundaries = false;
                            break;
                        }
                        cp = nextCp;
                    }

                    if (consecutiveErrors >= 10)
                    {
                        _logger.LogWarning("Too many consecutive errors in paragraph detection. Falling back to simple extraction.");
                        useParagraphBoundaries = false;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Failed to use paragraph boundaries: {ex.Message}. Falling back to simple method.");
                    useParagraphBoundaries = false;
                }
            }

            // Fallback: simple text extraction without paragraph boundaries
            if (!useParagraphBoundaries || defaultSection.Paragraphs.Count == 0)
            {
                defaultSection.Paragraphs.Clear();
                var currentParagraph = new WvWareNet.Core.Paragraph();
                defaultSection.Paragraphs.Add(currentParagraph);

                // Extract all text from document
                string fullText = pieceTable.GetTextForRange(0, (int)fib.CcpText, wordDocMs);

                _logger.LogInfo("[FALLBACK] Starting fallback paragraph split...");
                // Split by paragraph markers
                string[] parts = fullText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                _logger.LogInfo($"[FALLBACK] Split into {parts.Length} parts. Creating paragraphs...");
                foreach (var part in parts)
                {
                    if (!string.IsNullOrEmpty(part))
                    {
                        var paragraph = new WvWareNet.Core.Paragraph();
                        var run = new WvWareNet.Core.Run { Text = part };
                        paragraph.Runs.Add(run);
                        defaultSection.Paragraphs.Add(paragraph);
                    }
                }
                _logger.LogInfo($"[FALLBACK] Finished creating paragraphs. Total: {defaultSection.Paragraphs.Count}");

                // Remove empty paragraph at the beginning if it exists
                if (defaultSection.Paragraphs.Count > 0 && defaultSection.Paragraphs[0].Runs.Count == 0)
                {
                    defaultSection.Paragraphs.RemoveAt(0);
                }
            }

            // Add default headers/footers/footnotes
            _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            _documentModel.Footers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote { ReferenceId = 1 });

            // Log document structure for debugging
            _logger.LogInfo($"[STRUCTURE] Document has {_documentModel.Sections.Count} section(s)");
            for (int i = 0; i < _documentModel.Sections.Count; i++)
            {
                var section = _documentModel.Sections[i];
                _logger.LogInfo($"[STRUCTURE] Section {i + 1}: {section.Paragraphs.Count} paragraph(s)");
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
                    // Add paragraph separator
                    if (paragraph.Runs.Count > 0)
                        textBuilder.AppendLine();
                }
            }

            _logger.LogInfo($"[EXTRACTION] Extracted {runCount} runs with {charCount} characters");
            return textBuilder.ToString();
        }
    }
}
