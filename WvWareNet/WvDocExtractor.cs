using System.IO;
using System.Text;
using WvWareNet.Parsers;
using WvWareNet.Utilities;

namespace WvWareNet;

public class WvDocExtractor
{
    private readonly ILogger _logger;

    static WvDocExtractor()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public WvDocExtractor(ILogger logger)
    {
        _logger = logger;
    }

    public string ExtractText(string filePath, string password = null)
    {
        if (!File.Exists(filePath))
        {
            _logger.LogError($"File not found: {filePath}");
            throw new FileNotFoundException($"File not found: {filePath}");
        }

        byte[] fileData = File.ReadAllBytes(filePath);

        // Check for DOCX/OOXML (ZIP) signature
        if (fileData.Length >= 4 &&
            fileData[0] == 0x50 && fileData[1] == 0x4B && fileData[2] == 0x03 && fileData[3] == 0x04)
        {
            _logger.LogError("File appears to be a Word 2007+ (.docx) file with the wrong extension.");
            throw new InvalidDataException("File appears to be a Word 2007+ (.docx) file with the wrong extension.");
        }

        // Handle Word95 files using magic bytes
        bool isWord95 = fileData.Length > 0x200 && fileData[0x200] == 0xEC && fileData[0x201] == 0xA5;
        if (isWord95)
        {
            _logger.LogInfo("Detected Word95 document format");
            // Check if file is encrypted
            bool isEncrypted = fileData.Length > 0x20 &&
                             (fileData[0x0B] & 0x01) == 0x01;
            _logger.LogInfo($"Word95 document - Encrypted: {isEncrypted}");

            if (isEncrypted)
            {
                try
                {
                    uint lKey = BitConverter.ToUInt32(fileData, 0x00);
                    fileData = Word95Decryptor.Decrypt(fileData, password ?? string.Empty, lKey);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Decryption failed: {ex.Message}");
                    throw new InvalidDataException("Failed to decrypt document. Invalid password or corrupted file.");
                }
            }
        }

        // Always use CFBF parser for all files (including Word95)
        using var fileStream = new MemoryStream(fileData);
        var cfbfParser = new CompoundFileBinaryFormatParser(fileStream);

        // Try to parse CFBF header, fallback if invalid
        try
        {
            cfbfParser.ParseHeader();
        }
        catch (InvalidDataException ex)
        {
            _logger.LogError($"CFBF/OLE header invalid: {ex.Message}");
            _logger.LogInfo("Attempting fallback extraction for non-OLE/CFBF .doc file");
            return ExtractTextFallback(fileData);
        }

        // Find streams by name
        var entries = cfbfParser.ParseDirectoryEntries();
        var wordDocEntry = entries.Find(e => e.Name.Contains("WordDocument", StringComparison.OrdinalIgnoreCase));
        if (wordDocEntry == null)
            throw new InvalidDataException("WordDocument stream not found in CFBF file.");
        byte[] wordDocumentStream = cfbfParser.ReadStream(wordDocEntry);
        var fib = Core.FileInformationBlock.Parse(wordDocumentStream);

        // Try to get the Table stream and CLX
        var tableEntry = entries.Find(e => e.Name.Contains("1Table", StringComparison.OrdinalIgnoreCase))
            ?? entries.Find(e => e.Name.Contains("0Table", StringComparison.OrdinalIgnoreCase));
        byte[] tableStream = null;
        if (tableEntry != null)
            tableStream = cfbfParser.ReadStream(tableEntry);
        byte[] clx = null;
        if (tableStream != null && fib.FcClx >= 0 && fib.LcbClx > 0 && fib.FcClx + fib.LcbClx <= tableStream.Length)
        {
            clx = new byte[fib.LcbClx];
            Array.Copy(tableStream, fib.FcClx, clx, 0, fib.LcbClx);
        }

        // Always use WordDocumentParser for Word97+ files, regardless of CLX/piece table presence
        string extractedText = null;
        try
        {
            var wordDocumentParser = new WordDocumentParser(cfbfParser, _logger);
            wordDocumentParser.ParseDocument(password);
            extractedText = wordDocumentParser.ExtractText();
        }
        catch(NotSupportedException ex)
        {
            extractedText = ex.Message;
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error parsing document with CFBF: {ex.Message}");
            _logger.LogInfo("Attempting fallback text extraction");
            extractedText = ExtractTextFallback(fileData);
        }

        _logger.LogInfo($"[Extractor] Text after WordDocumentParser: '{extractedText?.Substring(0, Math.Min(50, extractedText.Length)) ?? "null"}'");

        // If extracted text is empty, try fallback anyway
        if (string.IsNullOrWhiteSpace(extractedText))
        {
            _logger.LogInfo("[Extractor] Extracted text was empty, attempting fallback");
            extractedText = ExtractTextFallback(fileData);
            _logger.LogInfo($"[Extractor] Text after Fallback: '{extractedText?.Substring(0, Math.Min(50, extractedText.Length)) ?? "null"}'");
        }

        return extractedText;
    }

    private string ExtractTextFallback(byte[] fileData)
    {
        string extractedText = null;

        // Try to parse as flat Word binary doc (FIB at start)
        bool triedFib = false;
        try
        {
            var fib = Core.FileInformationBlock.Parse(fileData);
            triedFib = true;
            // Known Word versions: 0x0062, 0x0063, 0x0065, 0x0076, 0x00C1, 0x00D9, 0x0101, 0x0112
            ushort[] knownFibVersions = { 0x0062, 0x0063, 0x0065, 0x0076, 0x00C1, 0x00D9, 0x0101, 0x0112 };
            if (Array.Exists(knownFibVersions, v => v == fib.NFib) && fib.FcMin < fib.FcMac && fib.FcMac <= fileData.Length)
            {
                int textStart = (int)fib.FcMin;
                int textLen = (int)(fib.FcMac - fib.FcMin);
                if (textStart >= 0 && textLen > 0 && textStart + textLen <= fileData.Length)
                {
                    byte[] textBytes = new byte[textLen];
                    Array.Copy(fileData, textStart, textBytes, 0, textLen);

                    // Try UTF-16LE first, then fallback to Windows-1252, then ASCII
                    try
                    {
                        extractedText = Encoding.Unicode.GetString(textBytes);
                        if (!string.IsNullOrWhiteSpace(extractedText))
                        {
                            _logger.LogInfo("Flat .doc: Successfully decoded main text as UTF-16LE");
                        }
                    }
                    catch { }

                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        try
                        {
                            extractedText = Encoding.GetEncoding(1251).GetString(textBytes);
                            if (!string.IsNullOrWhiteSpace(extractedText))
                            {
                                _logger.LogInfo("Flat .doc: Successfully decoded main text as Windows-1251 (Cyrillic)");
                            }
                        }
                        catch { }
                    }

                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        try
                        {
                            extractedText = Encoding.GetEncoding(1252).GetString(textBytes);
                            if (!string.IsNullOrWhiteSpace(extractedText))
                            {
                                _logger.LogInfo("Flat .doc: Successfully decoded main text as Windows-1252");
                            }
                        }
                        catch { }
                    }

                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        var sb = new StringBuilder();
                        foreach (byte b in textBytes)
                        {
                            if (b >= 32 && b <= 126)
                                sb.Append((char)b);
                        }
                        extractedText = sb.ToString();
                        _logger.LogInfo("Flat .doc: Used ASCII extraction for main text");
                    }
                }
            }
        }
        catch
        {
            // Ignore and fallback to generic extraction
        }

        // If still empty, use sliding window to extract readable text sequences
        if (string.IsNullOrWhiteSpace(extractedText))
        {
            _logger.LogInfo("Sliding window: extracting readable text sequences from file");
            var sb = new StringBuilder();
            int minSeqLen = 8;

            // ASCII sequences
            int asciiCount = 0;
            var asciiSeq = new StringBuilder();
            for (int i = 0; i < fileData.Length; i++)
            {
                byte b = fileData[i];
                if (b >= 32 && b <= 126)
                {
                    asciiSeq.Append((char)b);
                    asciiCount++;
                }
                else
                {
                    if (asciiCount >= minSeqLen)
                    {
                        sb.AppendLine(asciiSeq.ToString());
                    }
                    asciiSeq.Clear();
                    asciiCount = 0;
                }
            }
            if (asciiCount >= minSeqLen)
            {
                sb.AppendLine(asciiSeq.ToString());
            }

            // UTF-16LE sequences
            int unicodeCount = 0;
            var unicodeSeq = new StringBuilder();
            for (int i = 0; i < fileData.Length - 1; i += 2)
            {
                ushort ch = (ushort)(fileData[i] | (fileData[i + 1] << 8));
                char c = (char)ch;
                if (!char.IsControl(c) && !char.IsSurrogate(c) && (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c)))
                {
                    unicodeSeq.Append(c);
                    unicodeCount++;
                }
                else
                {
                    if (unicodeCount >= minSeqLen)
                    {
                        sb.AppendLine(unicodeSeq.ToString());
                    }
                    unicodeSeq.Clear();
                    unicodeCount = 0;
                }
            }
            if (unicodeCount >= minSeqLen)
            {
                sb.AppendLine(unicodeSeq.ToString());
            }

            extractedText = sb.ToString();
        }

        // If still empty, fallback to previous generic extraction
        if (string.IsNullOrWhiteSpace(extractedText))
        {
            // First try decoding as UTF-16LE (common in Word documents)
            try
            {
                extractedText = Encoding.Unicode.GetString(fileData);
                if (!string.IsNullOrWhiteSpace(extractedText))
                {
                    _logger.LogInfo("Fallback: Successfully decoded as UTF-16LE");
                }
            }
            catch
            {
                // Continue to other methods
            }

            // If still empty, try UTF-8
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                try
                {
                    extractedText = Encoding.UTF8.GetString(fileData);
                    if (!string.IsNullOrWhiteSpace(extractedText))
                    {
                        _logger.LogInfo("Fallback: Successfully decoded as UTF-8");
                    }
                }
                catch
                {
                    // Continue to other methods
                }
            }

            // If still empty, extract printable ASCII characters
            if (string.IsNullOrWhiteSpace(extractedText))
            {
                _logger.LogInfo("Fallback: Using ASCII extraction");
                var sb = new StringBuilder();
                foreach (byte b in fileData)
                {
                    if (b >= 32 && b <= 126)  // Printable ASCII range
                    {
                        sb.Append((char)b);
                    }
                }
                extractedText = sb.ToString();
            }

            // Clean up extracted text - remove non-printable characters
            var cleanText = new StringBuilder();
            foreach (char c in extractedText)
            {
                if (char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c))
                {
                    cleanText.Append(c);
                }
            }

            extractedText = cleanText.ToString();
        }

        return extractedText;
    }
}
