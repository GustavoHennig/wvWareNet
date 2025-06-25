using WvWareNet.Utilities;

namespace WvWareNet.Tests;

public class NullLogger : ILogger
{
    public void LogDebug(string message) { }
    public void LogInfo(string message) { }
    public void LogWarning(string message) { }
    public void LogError(string message) { }
    public void LogError(string message, Exception exception) { }
}
