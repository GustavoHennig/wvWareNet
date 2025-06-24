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
            var currentParagraph = new WvWareNet.Core.Paragraph();
            defaultSection.Paragraphs.Add(currentParagraph);

            for (int i = 0; i < pieceTable.Pieces.Count; i++)
            {
                string text = pieceTable.GetTextForPiece(i, wordDocMs);

                // Convert CHPX to CharacterProperties
                var chpx = pieceTable.Pieces[i].Chpx;
                var charProps = new WvWareNet.Core.CharacterProperties();
                if (chpx != null && chpx.Length > 0)
                {
                    // Minimal CHPX parsing: look for known sprm codes for bold, italic, underline
                    // This is a simplified example; real CHPX parsing is more complex
                    for (int b = 0; b < chpx.Length - 2; b++)
                    {
                        byte sprm = chpx[b];
                        byte val = chpx[b + 2]; // skip operand size byte
                        switch (sprm)
                        {
                            case 0x08: // sprmCFBold
                                charProps.IsBold = val != 0;
                                break;
                            case 0x09: // sprmCFItalic
                                charProps.IsItalic = val != 0;
                                break;
                            case 0x0A: // sprmCFStrike
                                charProps.IsStrikeThrough = val != 0;
                                break;
                            case 0x0B: // sprmCFOutline
                                // Not mapped
                                break;
                            case 0x0C: // sprmCFShadow
                                // Not mapped
                                break;
                            case 0x0D: // sprmCFSmallCaps
                                charProps.IsSmallCaps = val != 0;
                                break;
                            case 0x0E: // sprmCFCaps
                                charProps.IsAllCaps = val != 0;
                                break;
                            case 0x0F: // sprmCFVanish
                                charProps.IsHidden = val != 0;
                                break;
                            case 0x18: // sprmCFU (underline)
                                charProps.IsUnderlined = val != 0;
                                break;
                        }
                    }
                }

                var run = new WvWareNet.Core.Run { Text = text, Properties = charProps };
                currentParagraph.Runs.Add(run);

                // Simple heuristic for new paragraphs: if a piece ends with a newline, start a new paragraph.
                // This is a placeholder and needs proper paragraph boundary detection from Word's binary format.
                if (text.EndsWith("\r\n") || text.EndsWith("\n") || text.EndsWith("\r"))
                {
                    currentParagraph = new WvWareNet.Core.Paragraph();
                    defaultSection.Paragraphs.Add(currentParagraph);
                }
            }

            // Remove the last empty paragraph if it was created due to a trailing newline
            if (currentParagraph.Runs.Count == 0 && defaultSection.Paragraphs.Count > 1)
            {
                defaultSection.Paragraphs.RemoveAt(defaultSection.Paragraphs.Count - 1);
            }

            // --- Header/Footer/Footnote Extraction (simplified) ---
            // NOTE: This is a placeholder. Real implementation would parse PLCF structures for headers/footers/footnotes.
            // For demonstration, add empty header/footer/footnote objects.
            _documentModel.Headers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            _documentModel.Footers.Add(new WvWareNet.Core.HeaderFooter { Type = WvWareNet.Core.HeaderFooterType.Default });
            _documentModel.Footnotes.Add(new WvWareNet.Core.Footnote { ReferenceId = 1 });
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
