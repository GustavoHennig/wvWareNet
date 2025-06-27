using System;
using System.IO;
using WvWareNet;
using WvWareNet.Utilities;

namespace WvWareNetConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WvWareNetConsole <path_to_doc_file> [output_file] [-p password]");
                return;
            }

            // Parse arguments
            string filePath = null;
            string outputPath = null;
            string password = null;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-p" && i + 1 < args.Length)
                {
                    password = args[i + 1];
                    i++; // Skip password value
                }
                else if (filePath == null)
                {
                    filePath = args[i];
                }
                else if (outputPath == null)
                {
                    outputPath = args[i];
                }
            }

            if (filePath == null)
            {
                Console.WriteLine("Error: Input file is required.");
                Console.WriteLine("Usage: WvWareNetConsole <path_to_doc_file> [output_file] [-p password]");
                return;
            }

            if (outputPath == null)
            {
                outputPath = Path.ChangeExtension(filePath, ".txt");
            }

            var logger = new ConsoleLogger();
            var extractor = new WvDocExtractor(logger);

            try
            {
                string extractedText;

                // Check if this might be a Word95 file
                bool isWord95 = filePath.EndsWith(".word95.doc", StringComparison.OrdinalIgnoreCase);
                if (isWord95 && string.IsNullOrEmpty(password))
                {
                    byte[] fileData = File.ReadAllBytes(filePath);
                    bool isEncrypted = fileData.Length > 0x20 && (fileData[0x0B] & 0x01) == 0x01;

                    if (isEncrypted)
                    {
                        Console.Write("This is an encrypted Word95 document. Enter password: ");
                        password = Console.ReadLine();
                    }
                }

                extractedText = extractor.ExtractText(filePath, password);
                File.WriteAllText(outputPath, extractedText);
                Console.WriteLine($"Text extracted to: {outputPath}");
            }
            catch (Exception ex)
            {
                logger.LogError($"An error occurred during text extraction: {ex.Message}", ex);

                // If it's a Word95 decryption error, suggest trying with password
                if (ex.Message.Contains("Word95 decryption failed"))
                {
                    Console.WriteLine("Note: This appears to be an encrypted Word95 document.");
                    Console.WriteLine("Try running again with a password: WvWareNetConsole <file> <password>");
                }
            }
        }
    }
}
