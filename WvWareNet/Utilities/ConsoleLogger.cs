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
            Console.WriteLine($"[DEBUG] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
    }

    public void LogInfo(string message)
    {
        if (_minLogLevel <= LogLevel.Info)
            Console.WriteLine($"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
    }

    public void LogWarning(string message)
    {
        if (_minLogLevel <= LogLevel.Warning)
            Console.WriteLine($"[WARN] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
    }

    public void LogError(string message)
    {
        if (_minLogLevel <= LogLevel.Error)
            Console.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
    }

    public void LogError(string message, Exception exception)
    {
        if (_minLogLevel <= LogLevel.Error)
        {
            Console.WriteLine($"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}");
            Console.WriteLine($"Exception Type: {exception.GetType().FullName}");
            Console.WriteLine($"Exception Message: {exception.Message}");
            Console.WriteLine("Stack Trace:");
            Console.WriteLine(exception.StackTrace);
            if (exception.InnerException != null)
            {
                Console.WriteLine("Inner Exception:");
                Console.WriteLine($"Type: {exception.InnerException.GetType().FullName}");
                Console.WriteLine($"Message: {exception.InnerException.Message}");
            }
        }
    }
}
