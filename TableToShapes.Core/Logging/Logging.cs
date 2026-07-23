using System;

namespace TableToShapes.Core.Logging
{
    /// <summary>Severity levels, in ascending order. <see cref="None"/> disables all output.</summary>
    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warning = 2,
        Error = 3,
        None = 4
    }

    /// <summary>
    /// Minimal logging abstraction so the conversion pipeline can record what it does without
    /// depending on any concrete sink. Inject a <see cref="FileLogger"/> in production or
    /// <see cref="NullLogger"/> (the default) in tests.
    /// </summary>
    public interface ILogger
    {
        /// <summary>True if a message at this level would be written (lets callers skip building expensive messages).</summary>
        bool IsEnabled(LogLevel level);

        void Log(LogLevel level, string message, Exception exception = null);
    }

    /// <summary>Convenience level-specific methods over <see cref="ILogger"/>.</summary>
    public static class LoggerExtensions
    {
        public static void Debug(this ILogger logger, string message) => logger.Log(LogLevel.Debug, message);
        public static void Info(this ILogger logger, string message) => logger.Log(LogLevel.Info, message);
        public static void Warning(this ILogger logger, string message, Exception exception = null) => logger.Log(LogLevel.Warning, message, exception);
        public static void Error(this ILogger logger, string message, Exception exception = null) => logger.Log(LogLevel.Error, message, exception);
    }

    /// <summary>A logger that discards everything. Used as the default so a logger is never null.</summary>
    public sealed class NullLogger : ILogger
    {
        public static readonly NullLogger Instance = new NullLogger();
        private NullLogger() { }

        public bool IsEnabled(LogLevel level) => false;
        public void Log(LogLevel level, string message, Exception exception = null) { }
    }

    /// <summary>Parses a configured <see cref="LogLevel"/> name (case-insensitive) with a fallback.</summary>
    public static class LogLevelParser
    {
        public static LogLevel Parse(string value, LogLevel fallback = LogLevel.Info)
        {
            if (!string.IsNullOrWhiteSpace(value) && Enum.TryParse(value.Trim(), ignoreCase: true, out LogLevel level))
                return level;
            return fallback;
        }
    }
}
