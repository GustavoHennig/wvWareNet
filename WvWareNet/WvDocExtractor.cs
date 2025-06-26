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
        
        // Handle Word95 files using magic bytes
        if (fileData.Length > 0x200 && fileData[0x200] == 0xEC && fileData[0x201] == 0xA5)
        {
            // Check if file is encrypted
            bool isEncrypted = fileData.Length > 0x20 && 
                             (fileData[0x0B] & 0x01) == 0x01;

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
        
        var wordDocumentParser = new WordDocumentParser(cfbfParser, _logger);
        wordDocumentParser.ParseDocument(password);

        return wordDocumentParser.ExtractText();
    }
}
