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

                logger.LogInfo($"Starting text extraction for: {filePath}");
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                
                extractedText = extractor.ExtractText(filePath, password);
                
                stopwatch.Stop();
                logger.LogInfo($"Extraction completed in {stopwatch.ElapsedMilliseconds}ms");
                logger.LogInfo($"Extracted text length: {extractedText?.Length ?? 0}");
                
                if (!string.IsNullOrEmpty(extractedText))
                {
                    File.WriteAllText(outputPath, extractedText);
                    Console.WriteLine($"Text extracted to: {outputPath}");
                    Console.WriteLine("Extracted text preview:");
                    Console.WriteLine(extractedText.Substring(0, Math.Min(200, extractedText.Length)));
                }
                else
                {
                    logger.LogError("Extraction returned empty text. Please check the logs for details.");
                    Console.WriteLine("Extraction completed but returned empty text. Please check the logs for details.");
                }
            }
            catch (Exception ex)
            {
                // Directly output exception details to console
                Console.WriteLine($"An error occurred during text extraction: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine("Stack Trace:");
                Console.WriteLine(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    Console.WriteLine("Inner Exception:");
                    Console.WriteLine($"Type: {ex.InnerException.GetType().FullName}");
                    Console.WriteLine($"Message: {ex.InnerException.Message}");
                }

                // Also log using logger
                logger.LogError($"An error occurred during text extraction: {ex.ToString()}", ex);
            }
        }
    }
}
