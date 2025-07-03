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

    public string ExtractText(string filePath, string password = null, bool extractHeadersFooters = false)
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

        // Post-processing: remove unwanted chars 0A (LF) and 0C (FF)
        extractedText = RemoveUnwantedChars(extractedText);

        return extractedText;
    }

    private string RemoveUnwantedChars(string input)
    {
        if (input == null)
            return null;
        // Remove LF (\n, 0x0A) and FF (\f, 0x0C)
        return input.Replace("\f", "");
    }

    private string ExtractTextFallback(byte[] fileData)
    {
        // We should avoid using fallback implementation
        _logger.LogWarning("Using fallback text extraction is discouraged");
        return string.Empty;
    }
}
