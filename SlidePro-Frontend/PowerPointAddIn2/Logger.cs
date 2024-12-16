using System;
using System.IO;
using System.Threading;

public static class Logger
{
    private static readonly object _lock = new object();
    private static string logFilePath = Path.Combine(
        "C:\\Users\\nafay\\PPT_output",
        $"Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

    static Logger()
    {
        try
        {
            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
        }
        catch (Exception ex)
        {
            // If logging fails, there's not much we can do.
            // Optionally, handle exceptions or fallback mechanisms here.
        }
    }

    public static void Log(string message)
    {
        try
        {
            string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";
            lock (_lock)
            {
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
        }
        catch
        {
            // Suppress any logging exceptions to avoid disrupting the main application
        }
    }

    public static void LogException(string context, Exception ex)
    {
        Log($"{context} - Exception: {ex.Message}\nStack Trace: {ex.StackTrace}");
    }

    public static string GetLogFilePath()
    {
        return logFilePath;
    }
}
