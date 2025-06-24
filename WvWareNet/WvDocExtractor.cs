using System.IO;
using WvWareNet.Parsers;
using WvWareNet.Utilities;

namespace WvWareNet;

public class WvDocExtractor
{
    private readonly ILogger _logger;

    public WvDocExtractor(ILogger logger)
    {
        _logger = logger;
    }

    public string ExtractText(string filePath)
    {
        if (!File.Exists(filePath))
        {
            _logger.LogError($"File not found: {filePath}");
            throw new FileNotFoundException($"File not found: {filePath}");
        }

        using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        var cfbfParser = new CompoundFileBinaryFormatParser(fileStream);
        cfbfParser.ParseHeader();

        var wordDocumentParser = new WordDocumentParser(cfbfParser);
        wordDocumentParser.ParseDocument();

        return wordDocumentParser.ExtractText();
    }
}
