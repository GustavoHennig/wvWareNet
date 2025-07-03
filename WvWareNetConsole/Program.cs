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
                Console.WriteLine("Usage: WvWareNetConsole <path_to_doc_file> [output_file] [-p password] [--extract-headers-footers]");
                Console.WriteLine("  --extract-headers-footers   Extract headers and footers text (default: NO)");
                return;
            }

            // Parse arguments
            string filePath = null;
            string outputPath = null;
            string password = null;
            bool extractHeadersFooters = false;
            string logFilePath = null;
            WvWareNet.Utilities.LogLevel logLevel = WvWareNet.Utilities.LogLevel.Debug;

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-p" && i + 1 < args.Length)
                {
                    password = args[i + 1];
                    i++; // Skip password value
                }
                else if (args[i] == "--extract-headers-footers")
                {
                    extractHeadersFooters = true;
                }
                else if (args[i] == "--verbosity" && i + 1 < args.Length)
                {
                    if (Enum.TryParse<WvWareNet.Utilities.LogLevel>(args[i + 1], true, out var parsedLevel))
                    {
                        logLevel = parsedLevel;
                        i++;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid verbosity level: {args[i + 1]}");
                        return;
                    }
                }
                else if (args[i] == "--log-to-file" && i + 1 < args.Length)
                {
                    logFilePath = args[i + 1];
                    i++;
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

            WvWareNet.Utilities.ILogger logger;
            if (!string.IsNullOrEmpty(logFilePath))
            {
                logger = new WvWareNet.Utilities.FileLogger(logFilePath, logLevel);
            }
            else
            {
                logger = new WvWareNet.Utilities.ConsoleLogger(logLevel);
            }
            var extractor = new WvDocExtractor(logger);

            try
            {
                string extractedText;

                logger.LogInfo($"Starting text extraction for: {filePath}");
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                
                extractedText = extractor.ExtractText(filePath, password, extractHeadersFooters);
                
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
        finally
        {
            // Dispose logger if needed
            if (logger is IDisposable disposableLogger)
            {
                disposableLogger.Dispose();
            }
        }
    }
}
}
