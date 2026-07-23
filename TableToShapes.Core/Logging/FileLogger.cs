using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace TableToShapes.Core.Logging
{
    /// <summary>
    /// Writes log lines to a file, filtered by a configured minimum <see cref="LogLevel"/>.
    /// Self-rotating (truncates past <c>maxBytes</c>) so it cannot grow without bound, and
    /// best-effort: a failure to write never propagates to the caller. Thread-safe.
    /// </summary>
    public sealed class FileLogger : ILogger
    {
        private readonly string _path;
        private readonly LogLevel _minLevel;
        private readonly long _maxBytes;
        private readonly object _gate = new object();

        public FileLogger(string path, LogLevel minLevel, long maxBytes = 512 * 1024)
        {
            _path = path;
            _minLevel = minLevel;
            _maxBytes = maxBytes;
        }

        public bool IsEnabled(LogLevel level) => level != LogLevel.None && level >= _minLevel;

        public void Log(LogLevel level, string message, Exception exception = null)
        {
            if (!IsEnabled(level)) return;
            try
            {
                lock (_gate)
                {
                    if (_maxBytes > 0 && File.Exists(_path) && new FileInfo(_path).Length > _maxBytes)
                        File.Delete(_path);

                    var sb = new StringBuilder()
                        .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture))
                        .Append(" [").Append(level.ToString().ToUpperInvariant()).Append("] ")
                        .Append(message);
                    if (exception != null)
                        sb.Append(Environment.NewLine).Append(exception);
                    sb.Append(Environment.NewLine);

                    File.AppendAllText(_path, sb.ToString());
                }
            }
            catch { /* logging must never take the caller down */ }
        }
    }
}
