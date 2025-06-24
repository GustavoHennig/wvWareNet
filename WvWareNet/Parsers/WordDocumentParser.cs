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

            // Extract text using PieceTable
            _documentModel = new WvWareNet.Core.DocumentModel();
            using var wordDocMs = new System.IO.MemoryStream(wordDocStream);
            for (int i = 0; i < pieceTable.Pieces.Count; i++)
            {
                string text = pieceTable.GetTextForPiece(i, wordDocMs);
                var run = new WvWareNet.Core.Run { Text = text };
                var para = new WvWareNet.Core.Paragraph();
                para.Runs.Add(run);
                _documentModel.Paragraphs.Add(para);
            }
        }

        public string ExtractText()
        {
            if (_documentModel == null)
                throw new InvalidOperationException("Document not parsed. Call ParseDocument() first.");

            var text = new System.Text.StringBuilder();
            foreach (var para in _documentModel.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    text.Append(run.Text);
                }
                text.AppendLine();
            }
            return text.ToString();
        }
    }
}
