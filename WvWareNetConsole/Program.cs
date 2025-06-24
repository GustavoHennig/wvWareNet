using System;
using WvWareNet;
using WvWareNet.Utilities;

namespace WvWareNetConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WvWareNetConsole <path_to_doc_file>");
                return;
            }

            string filePath = args[0];
            var logger = new ConsoleLogger();
            var extractor = new WvDocExtractor(logger);

            try
            {
                string extractedText = extractor.ExtractText(filePath);
                Console.WriteLine("--- Extracted Text ---");
                Console.WriteLine(extractedText);
                Console.WriteLine("--- End of Extraction ---");
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred during text extraction: {ex.Message}", ex);
            }
        }
    }
}
