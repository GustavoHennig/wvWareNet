using System;
using System.IO;
using System.Linq;
using Xunit;
using WvWareNet;
using System.Diagnostics;
using Xunit.Sdk;

namespace WvWareNet.Tests
{
    public class ExamplesDocTests
    {
        public static IEnumerable<object[]> DocFiles()
        {
            // Determine solution root and examples folder
            var baseDir = AppContext.BaseDirectory;
            var examplesDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples"));
            var examplesLocalDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "..", "examples-local"));

            if (!Directory.Exists(examplesDir))
                throw new DirectoryNotFoundException($"Examples directory not found: {examplesDir}");

            // Find all .doc files
            var docFiles = Directory.GetFiles(examplesDir, "*.doc").ToList();


            if (Directory.Exists(examplesLocalDir))
            {
                docFiles.AddRange(Directory.GetFiles(examplesLocalDir, "*.doc"));
            }

            return docFiles
                .Where(w => File.Exists(Path.ChangeExtension(  w, ".expected.txt")))
                .Select(doc => new object[] { doc, Path.ChangeExtension(doc, ".expected.txt") });
        }

        [Theory]
        [MemberData(nameof(DocFiles))]
        public void ExtractedText_EqualsExpectedFile(string docPath, string expectedPath)
        {
            if (!File.Exists(expectedPath))
            {
                Debug.Print($"Expected file not found: {expectedPath}");
                throw SkipException.ForSkip($"Expected file not found: {expectedPath}");
            }

            var extractor = new WvDocExtractor(new NullLogger());
            var result = NormalizeText(extractor.ExtractText(docPath));
            var expected = NormalizeText(File.ReadAllText(expectedPath));
            bool isEqual = string.Equals(result, expected, StringComparison.InvariantCultureIgnoreCase);
            Assert.Equal(expected,result, true, true, true, true);
            if(!isEqual)
            {
                Debug.Print($"Mismatch in {docPath}");
            }

        }
        /// <summary>
        /// Normalizes text by standardizing line breaks to '\n' and trimming trailing whitespace from each line.
        /// </summary>
        public static string NormalizeText(string text)
        {
            if (text == null) return null;
            // Replace CRLF and CR with LF
            var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            // Trim trailing whitespace from each line
            var lines = normalized.Split('\n').Select(line => line.TrimEnd());
            return string.Join("\n", lines);
        }


    }
}
