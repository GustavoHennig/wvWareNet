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
                var run = new WvWareNet.Core.Run { Text = text };
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
