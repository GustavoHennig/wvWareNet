using System;
using System.Collections.Generic;
using System.IO;
using WvWareNet;
using WvWareNet.Utilities;
using Xunit;

namespace WvWareNet.Tests;

public class DocParsingTests
{
    public static IEnumerable<object[]> DocFiles => new List<object[]>
    {
        new object[] { "a-idade-media.word95.doc" },
        new object[] { "file-sample_libreoffice.doc" },
        new object[] { "news-example.doc" },
        new object[] { "test.doc" },
        new object[] { "test2.doc" }
    };

    [Theory]
    [MemberData(nameof(DocFiles))]
    public void AllSampleDocsParse(string fileName)
    {
        string filePath = Path.Combine(AppContext.BaseDirectory, fileName);
        Assert.True(File.Exists(filePath), $"Sample file not found: {filePath}");

        var extractor = new WvDocExtractor(new ConsoleLogger());
        string text = extractor.ExtractText(filePath);
        Assert.False(string.IsNullOrWhiteSpace(text));
    }

    [Fact]
    public void ConvertAllDocsInExternalFolderToTextWithWvSuffix()
    {
        string folder = @"";
        var docFiles = Directory.GetFiles(folder, "*.doc", SearchOption.TopDirectoryOnly);

        var extractor = new WvDocExtractor(new ConsoleLogger());

        foreach (var docFile in docFiles)
        {
            string text = extractor.ExtractText(docFile);
            Assert.False(string.IsNullOrWhiteSpace(text), $"No text extracted from {docFile}");

            string outputPath = Path.Combine(
                folder,
                Path.GetFileNameWithoutExtension(docFile) + "_wv.txt"
            );
            File.WriteAllText(outputPath, text);
            Assert.True(File.Exists(outputPath), $"Output file not created: {outputPath}");
        }
    }
}
