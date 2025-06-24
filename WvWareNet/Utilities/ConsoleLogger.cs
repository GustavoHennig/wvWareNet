namespace WvWareNet.Utilities;

public class ConsoleLogger : ILogger
{
    private readonly LogLevel _minLogLevel;

    public ConsoleLogger(LogLevel minLogLevel = LogLevel.Info)
    {
        _minLogLevel = minLogLevel;
    }

    public void LogDebug(string message)
    {
        if (_minLogLevel <= LogLevel.Debug)
        {
            Console.WriteLine($"[DEBUG] {DateTime.Now:HH:mm:ss} - {message}");
        }
    }

    public void LogInfo(string message)
    {
        if (_minLogLevel <= LogLevel.Info)
        {
            Console.WriteLine($"[INFO] {DateTime.Now:HH:mm:ss} - {message}");
        }
    }

    public void LogWarning(string message)
    {
        if (_minLogLevel <= LogLevel.Warning)
        {
            Console.WriteLine($"[WARN] {DateTime.Now:HH:mm:ss} - {message}");
        }
    }

    public void LogError(string message)
    {
        if (_minLogLevel <= LogLevel.Error)
        {
            Console.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} - {message}");
        }
    }

    public void LogError(string message, Exception exception)
    {
        if (_minLogLevel <= LogLevel.Error)
        {
            Console.WriteLine($"[ERROR] {DateTime.Now:HH:mm:ss} - {message}");
            Console.WriteLine($"Exception: {exception.GetType().Name}: {exception.Message}");
            Console.WriteLine(exception.StackTrace);
        }
    }
}

public enum LogLevel
{
    Debug,
    Info,
    Warning,
    Error
}
