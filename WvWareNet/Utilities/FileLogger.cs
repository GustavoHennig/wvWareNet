using System;
using System.IO;

namespace WvWareNet.Utilities
{
    public class FileLogger : ILogger, IDisposable
    {
        private readonly LogLevel _minLogLevel;
        private readonly StreamWriter _writer;

        public FileLogger(string filePath, LogLevel minLogLevel = LogLevel.Info)
        {
            _minLogLevel = minLogLevel;
            _writer = new StreamWriter(filePath, append: true) { AutoFlush = true };
        }

        public void LogDebug(string message)
        {
            if (_minLogLevel <= LogLevel.Debug)
                WriteLog("DEBUG", message);
        }

        public void LogInfo(string message)
        {
            if (_minLogLevel <= LogLevel.Info)
                WriteLog("INFO", message);
        }

        public void LogWarning(string message)
        {
            if (_minLogLevel <= LogLevel.Warning)
                WriteLog("WARN", message);
        }

        public void LogError(string message)
        {
            if (_minLogLevel <= LogLevel.Error)
                WriteLog("ERROR", message);
        }

        public void LogError(string message, Exception exception)
        {
            if (_minLogLevel <= LogLevel.Error)
            {
                WriteLog("ERROR", message);
                WriteLog("ERROR", $"Exception Type: {exception.GetType().FullName}");
                WriteLog("ERROR", $"Exception Message: {exception.Message}");
                WriteLog("ERROR", "Stack Trace:");
                WriteLog("ERROR", exception.StackTrace);
                if (exception.InnerException != null)
                {
                    WriteLog("ERROR", "Inner Exception:");
                    WriteLog("ERROR", $"Type: {exception.InnerException.GetType().FullName}");
                    WriteLog("ERROR", $"Message: {exception.InnerException.Message}");
                }
            }
        }

        private void WriteLog(string level, string message)
        {
            _writer.WriteLine($"[{level}] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
        }

        public void Dispose()
        {
            _writer?.Dispose();
        }
    }
}
