using System;
using System.Collections.Generic;
using WvWareNet.Parsers;

namespace WvWareNet.Parsers
{
    public class WordDocumentParser
    {
        private readonly CompoundFileBinaryFormatParser _cfbfParser;

        public WordDocumentParser(CompoundFileBinaryFormatParser cfbfParser)
        {
            _cfbfParser = cfbfParser;
        }

        private WvWareNet.Core.DocumentModel _documentModel;

        public void ParseDocument()
        {
            // Parse directory entries
            var entries = _cfbfParser.ParseDirectoryEntries();

            // Locate WordDocument and Table streams
            var wordDocEntry = entries.Find(e => e.Name == "WordDocument");
            var tableEntry = entries.Find(e => e.Name == "1Table") ?? entries.Find(e => e.Name == "0Table");

            if (wordDocEntry == null || tableEntry == null)
                throw new InvalidDataException("Required streams not found in CFBF file.");

            // Read stream data
            var wordDocStream = _cfbfParser.ReadStream(wordDocEntry);
            var tableStream = _cfbfParser.ReadStream(tableEntry);

            // Parse FileInformationBlock (FIB)
            var fib = WvWareNet.Core.FileInformationBlock.Parse(wordDocStream);

            // Extract CLX (piece table) data from Table stream using FIB offsets
            if (fib.FcClx < 0 || fib.LcbClx == 0 || fib.FcClx + fib.LcbClx > tableStream.Length)
                throw new InvalidDataException("Invalid CLX offsets in FIB.");

            byte[] clxData = new byte[fib.LcbClx];
            Array.Copy(tableStream, fib.FcClx, clxData, 0, fib.LcbClx);

            // Create PieceTable and parse
            var logger = new WvWareNet.Utilities.ConsoleLogger();
            var pieceTable = new WvWareNet.Core.PieceTable(logger);
            pieceTable.Parse(clxData);

            // Extract CHPX (character formatting) data from Table stream using PLCFCHPX
            // PLCFCHPX location is in FIB: FcPlcfbteChpx, LcbPlcfbteChpx
            var chpxList = new List<byte[]>();
            if (fib.FcPlcfbteChpx > 0 && fib.LcbPlcfbteChpx > 0 && fib.FcPlcfbteChpx + fib.LcbPlcfbteChpx <= tableStream.Length)
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
            if (fib.FcPlcfbtePapx > 0 && fib.LcbPlcfbtePapx > 0 && fib.FcPlcfbtePapx + fib.LcbPlcfbtePapx <= tableStream.Length)
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
            if (fib.FcPlcfhdd > 0 && fib.LcbPlcfhdd > 0 && fib.FcPlcfhdd + fib.LcbPlcfhdd <= tableStream.Length)
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
                    // For demonstration, just create a header/footer with the FC as a placeholder
                    _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter
                    {
                        Type = WvWareNet.Core.HeaderFooterType.Default,
                        Paragraphs = { new WvWareNet.Core.Paragraph { Runs = { new WvWareNet.Core.Run { Text = $"Header/Footer at FC {fc}" } } } }
                    });
                }
            }
            else
            {
                _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            }

            // Parse PLCF for footnotes if available
            if (fib.FcPlcfftn > 0 && fib.LcbPlcfftn > 0 && fib.FcPlcfftn + fib.LcbPlcfftn <= tableStream.Length)
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
                    _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote
                    {
                        ReferenceId = i + 1,
                        Paragraphs = { new WvWareNet.Core.Paragraph { Runs = { new WvWareNet.Core.Run { Text = $"Footnote at FC {fc}" } } } }
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
                // Optionally add section breaks or properties here
                // textBuilder.AppendLine("--- SECTION ---"); 

                foreach (var paragraph in section.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        textBuilder.Append(run.Text);
                    }
                    // Add a newline after each paragraph, unless it's the last run already contains one
                    if (paragraph.Runs.Count > 0 && !string.IsNullOrEmpty(paragraph.Runs[^1].Text) && 
                        !(paragraph.Runs[^1].Text.EndsWith("\r\n") || paragraph.Runs[^1].Text.EndsWith("\n") || paragraph.Runs[^1].Text.EndsWith("\r")))
                    {
                        textBuilder.AppendLine();
                    }
                }
            }
            return textBuilder.ToString();
        }
    }
}
