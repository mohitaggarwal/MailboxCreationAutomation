using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DPClientSDK
{
    public class FileLogger
    {
        private const string FILE_EXT = ".log";
        private readonly string datetimeFormat;
        private readonly string logFilename;
        private readonly string errorLogFilename;

        /// <summary>
        /// Initiate an instance of FileLogger class constructor.
        /// If log file does not exist, it will be created automatically.
        /// </summary>
        public FileLogger()
        {
            datetimeFormat = "yyyy-MM-dd HH:mm:ss.fff";
            string logFolderName = "Logs";
            if (!Directory.Exists(logFolderName))
            {
                Directory.CreateDirectory(logFolderName);
            }
            logFilename = Path.Combine(logFolderName,
                            $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss-fff")}{FILE_EXT}");
            errorLogFilename = Path.Combine(logFolderName, "error.log");
            // Log file header line
            string logHeader = logFilename + " is created.";
            if (!System.IO.File.Exists(logFilename))
            {
                WriteLine(System.DateTime.Now.ToString(datetimeFormat) + " " + logHeader, false);
            }
        }

        /// <summary>
        /// Log a DEBUG message
        /// </summary>
        /// <param name="text">Message</param>
        public void Debug(string text)
        {
            WriteFormattedLog(LogLevel.DEBUG, text);
        }

        /// <summary>
        /// Log an ERROR message
        /// </summary>
        /// <param name="text">Message</param>
        public void Error(string text)
        {
            WriteFormattedLog(LogLevel.ERROR, text);
            WriteError(text);
        }

        /// <summary>
        /// Log a FATAL ERROR message
        /// </summary>
        /// <param name="text">Message</param>
        public void Fatal(string text)
        {
            WriteFormattedLog(LogLevel.FATAL, text);
        }

        /// <summary>
        /// Log an INFO message
        /// </summary>
        /// <param name="text">Message</param>
        public void Info(string text)
        {
            WriteFormattedLog(LogLevel.INFO, text);
        }

        /// <summary>
        /// Log a TRACE message
        /// </summary>
        /// <param name="text">Message</param>
        public void Trace(string text)
        {
            WriteFormattedLog(LogLevel.TRACE, text);
        }

        /// <summary>
        /// Log a WARNING message
        /// </summary>
        /// <param name="text">Message</param>
        public void Warning(string text)
        {
            WriteFormattedLog(LogLevel.WARNING, text);
        }

        private void WriteLine(string text, bool append = true)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilename, append, System.Text.Encoding.UTF8))
                {
                    if (!string.IsNullOrEmpty(text))
                    {
                        writer.WriteLine(text);
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private void WriteError(string text, bool append = true)
        {
            try
            {
                string errorMessage = $"{DateTime.Now.ToString(datetimeFormat)} [ERROR]: {text}";
                using (StreamWriter writer = new StreamWriter(errorLogFilename, append, System.Text.Encoding.UTF8))
                {
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        writer.WriteLine(errorMessage);
                    }
                }
            }
            catch
            {
                throw;
            }
        }

        private void WriteFormattedLog(LogLevel level, string text)
        {
            string pretext;
            switch (level)
            {
                case LogLevel.TRACE:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [TRACE]   ";
                    break;
                case LogLevel.INFO:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [INFO]    ";
                    break;
                case LogLevel.DEBUG:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [DEBUG]   ";
                    break;
                case LogLevel.WARNING:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [WARNING]   ";
                    break;
                case LogLevel.ERROR:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [ERROR]   ";
                    break;
                case LogLevel.FATAL:
                    pretext = System.DateTime.Now.ToString(datetimeFormat) + " [FATAL]   ";
                    break;
                default:
                    pretext = "";
                    break;
            }

            WriteLine(pretext + text);
        }

        [System.Flags]
        private enum LogLevel
        {
            TRACE,
            INFO,
            DEBUG,
            WARNING,
            ERROR,
            FATAL
        }
    }

    public static class Logger
    {
        private static FileLogger fileLogger;

        public static FileLogger FileLogger
        {
            get
            {
                if (fileLogger == null)
                {
                    fileLogger = new FileLogger();
                }
                return fileLogger;
            }
        }
    }
}
