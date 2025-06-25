using System.IO;
using WvWareNet.Parsers;
using WvWareNet.Utilities;
using System.Linq;

namespace WvWareNet;

public class WvDocExtractor
{
    private readonly ILogger _logger;

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
        
        // Check if this is a CFBF (OLE) file by magic bytes
        bool isOleCf = fileData.Length > 8 &&
                       fileData[0] == 0xD0 && fileData[1] == 0xCF &&
                       fileData[2] == 0x11 && fileData[3] == 0xE0 &&
                       fileData[4] == 0xA1 && fileData[5] == 0xB1 &&
                       fileData[6] == 0x1A && fileData[7] == 0xE1;

        // Check if this is a Word95 file (by extension or magic bytes)
        bool isWord95 = filePath.EndsWith(".word95.doc", StringComparison.OrdinalIgnoreCase) ||
                       (fileData.Length > 0x200 && fileData[0x200] == 0xEC && fileData[0x201] == 0xA5);

        if (isWord95 && !isOleCf)
        {
            // Check if file is actually encrypted by reading FIB flags
            bool isEncrypted = fileData.Length > 0x20 && 
                             (fileData[0x0B] & 0x01) == 0x01; // Check FIB.fEncrypted flag

            if (isEncrypted)
            {
                try 
                {
                    // Read FIB to get lKey for decryption
                    uint lKey = BitConverter.ToUInt32(fileData, 0x00);
                    fileData = Word95Decryptor.Decrypt(fileData, password ?? string.Empty, lKey);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Word95 decryption failed: {ex.Message}");
                    throw new InvalidDataException("Failed to decrypt Word95 document. Invalid password or corrupted file.");
                }
            }
            else
            {
                _logger.LogInfo("Processing non-encrypted Word95 document");
            }

            // Directly parse as flat file (not CFBF)
            var wordDocumentParser = new WordDocumentParser(null);
            wordDocumentParser.ParseWord95FlatFile(fileData);
            return wordDocumentParser.ExtractText();
        }

        // Default: treat as CFBF (OLE) file
        using var fileStream = new MemoryStream(fileData);
        var cfbfParser = new CompoundFileBinaryFormatParser(fileStream);
        
        var wordDocumentParser2 = new WordDocumentParser(cfbfParser);
        wordDocumentParser2.ParseDocument();

        return wordDocumentParser2.ExtractText();
    }
}
