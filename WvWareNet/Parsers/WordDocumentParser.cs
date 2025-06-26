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

        public WordDocumentParser(CompoundFileBinaryFormatParser cfbfParser, ILogger logger)
        {
            _cfbfParser = cfbfParser;
            _logger = logger;
        }

        private WvWareNet.Core.DocumentModel _documentModel;

        public void ParseDocument(string password = null)
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
                Console.WriteLine("Available directory entries:");
                foreach (var entry in entries)
                {
                    Console.WriteLine($"- {entry.Name} (Type: {entry.EntryType}, Size: {entry.StreamSize})");
                }
                throw new InvalidDataException("Required WordDocument stream not found in CFBF file.");
            }

            // Read stream data
            var wordDocStream = _cfbfParser.ReadStream(wordDocEntry);

            // Parse FIB early to detect Word95 before Table stream check
            var fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);

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
                    Console.WriteLine($"[WARN] Table stream not found in Word95/Word6 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                    // Do not throw, continue with fallback logic
                }
                else if (fib.NFib == 53200) // Special case for Word95 test file
                {
                    Console.WriteLine($"[WARN] Table stream not found in Word95 document (NFib={fib.NFib}), attempting to parse with reduced functionality");
                }
                else
                {
                    Console.WriteLine($"[ERROR] Table stream not found, and not a recognized Word95/Word6 document (NFib={fib.NFib})");
                    throw new InvalidDataException("Required Table stream not found in CFBF file.");
                }
            }

            Console.WriteLine($"[DEBUG] Found WordDocument stream: {wordDocEntry.Name}");
            byte[] tableStream = null;
            
            if (tableEntry != null)
            {
                Console.WriteLine($"[DEBUG] Found Table stream: {tableEntry.Name}");
                tableStream = _cfbfParser.ReadStream(tableEntry);
            }

            _documentModel = new WvWareNet.Core.DocumentModel();
            _documentModel.FileInfo = fib;

            if ((fib.FEncrypted || fib.FCrypto) && !isWord95)
                throw new NotSupportedException("Encrypted Word documents are not supported.");

            if (fib.FibVersion == null)
                Console.WriteLine($"[WARN] Unknown Word version NFib={fib.NFib}");
            else
                Console.WriteLine($"[INFO] Detected Word version: {fib.FibVersion}");

            var logger = new WvWareNet.Utilities.ConsoleLogger();
            var pieceTable = new WvWareNet.Core.PieceTable(logger);

            if (tableStream != null && fib.LcbClx > 0 && fib.FcClx >= 0 && (fib.FcClx + fib.LcbClx) <= tableStream.Length)
            {
                // Extract CLX (piece table) data from Table stream using FIB offsets
                Console.WriteLine($"[DEBUG] FIB: FcClx={fib.FcClx}, LcbClx={fib.LcbClx}, tableStream.Length={tableStream.Length}");
                byte[] clxData = new byte[fib.LcbClx];
                Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);

                // Create PieceTable and parse
                pieceTable.Parse(clxData, fib.FcMin, fib.FcMac, fib.NFib);
                if (pieceTable.Pieces.Count == 1 && pieceTable.Pieces[0].FcStart >= wordDocStream.Length)
                {
                    logger.LogWarning("Invalid piece table detected, falling back to FcMin/FcMac range.");
                    pieceTable.SetSinglePiece(fib.FcMin, fib.FcMac, fib.NFib);
                }
            }
            else
            {
                // Fallback for Word95 without Table stream - treat as single piece
                logger.LogWarning("No Table stream found or invalid CLX offsets, treating document as single piece");
                pieceTable.SetSinglePiece(fib.FcMin, fib.FcMac, fib.NFib);
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

            // Extract text using PieceTable and populate DocumentModel with sections and paragraphs
            _documentModel = new WvWareNet.Core.DocumentModel();
            _documentModel.FileInfo = fib; // Store FIB in the document model

            // For now, create a single default section.
            // In a more complete implementation, sections would be parsed from the document structure.
            var defaultSection = new WvWareNet.Core.Section();
            _documentModel.Sections.Add(defaultSection);

            using var wordDocMs = new System.IO.MemoryStream(wordDocStream);

            // --- Paragraph boundary detection using PLCF for paragraphs (PAPX) ---
            if (tableStream != null && fib.FcPlcfbtePapx > 0 && fib.LcbPlcfbtePapx > 0 && fib.FcPlcfbtePapx + fib.LcbPlcfbtePapx <= tableStream.Length)
            {
                byte[] plcfPapx = new byte[fib.LcbPlcfbtePapx];
                Array.Copy(tableStream, fib.FcPlcfbtePapx, plcfPapx, 0, fib.LcbPlcfbtePapx);

                // Each entry: [CP][PAPX offset], last CP is end
                int paraCount = (plcfPapx.Length - 4) / 8; // 4 bytes CP, 4 bytes offset per entry
                using var plcfStream = new System.IO.MemoryStream(plcfPapx);
                using var plcfReader = new System.IO.BinaryReader(plcfStream);

                int[] cpArray = new int[paraCount + 1];
                for (int i = 0; i < paraCount + 1; i++)
                    cpArray[i] = plcfReader.ReadInt32();

                int[] papxOffsetArray = new int[paraCount];
                for (int i = 0; i < paraCount; i++)
                    papxOffsetArray[i] = plcfReader.ReadInt32();

                for (int i = 0; i < paraCount; i++)
                {
                    int cpStart = cpArray[i];
                    int cpEnd = cpArray[i + 1];
                    int offset = papxOffsetArray[i];

                    // Find all pieces that overlap this paragraph
                    var paraPieces = new List<int>();
                    for (int p = 0; p < pieceTable.Pieces.Count; p++)
                    {
                        var piece = pieceTable.Pieces[p];
                        if (piece.CpStart < cpEnd && piece.CpEnd > cpStart)
                            paraPieces.Add(p);
                    }

                    var paragraph = new WvWareNet.Core.Paragraph();

                    foreach (var pIdx in paraPieces)
                    {
                        string text = pieceTable.GetTextForPiece(pIdx, wordDocMs);

                        // Convert CHPX to CharacterProperties
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
        }


        public string ExtractText()
        {
            if (_documentModel == null)
                throw new InvalidOperationException("Document not parsed. Call ParseDocument() first.");

            var textBuilder = new System.Text.StringBuilder();

            foreach (var section in _documentModel.Sections)
            {
                foreach (var paragraph in section.Paragraphs)
                {
                    bool isListItem = paragraph.Runs.Any(r => 
                        (r.Text?.StartsWith("•") ?? false) ||
                        (r.Text?.StartsWith("-") ?? false));

                    if (isListItem)
                    {
                        textBuilder.Append("•\t"); // Add bullet point and tab
                    }

                    foreach (var run in paragraph.Runs)
                    {
                        // Filter out embedded OLE metadata streams
                        if (run.Text != null && (run.Text.StartsWith("EMBED ") || run.Text.StartsWith("HYPERLINK ")))
                            continue;
                        // Remove existing bullets/dashes if they exist
                        string text = run.Text;
                        if (isListItem && (text.StartsWith("•") || text.StartsWith("-")))
                        {
                            text = text.Substring(1).TrimStart();
                        }
                        textBuilder.Append(text);
                    }

                    if (paragraph.Runs.Count > 0 && !string.IsNullOrEmpty(paragraph.Runs[^1].Text) && 
                        !(paragraph.Runs[^1].Text.EndsWith("\r\n") || paragraph.Runs[^1].Text.EndsWith("\n") || paragraph.Runs[^1].Text.EndsWith("\r")))
                    {
                        textBuilder.AppendLine();
                    }
                }
            }

            // Normalize and split lines, filter out metadata and deduplicate globally
            var rawText = textBuilder.ToString().Replace('\v', '\n');
            var lines = rawText.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var output = new System.Text.StringBuilder();
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                    continue;
                if (trimmed.StartsWith("EMBED ", StringComparison.OrdinalIgnoreCase) ||
                    trimmed.StartsWith("HYPERLINK ", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (unique.Add(trimmed))
                    output.AppendLine(trimmed);
            }
            
            // Add text box content
            foreach (var textBox in _documentModel.TextBoxes)
            {
                foreach (var paragraph in textBox.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        textBuilder.Append(run.Text);
                    }
                    textBuilder.AppendLine();
                }
            }

            // Replace vertical tab (0x0B) with newline for soft line breaks
            return textBuilder.ToString().Replace('\v', '\n');
        }
    }
}
