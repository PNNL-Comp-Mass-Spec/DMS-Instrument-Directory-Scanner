using PRISM.Logging;
using PRISM;
using System;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Base class that implements logging methods
    /// </summary>
    public abstract class LoggerBase
    {
        /// <summary>
        /// Show a status message at the console and optionally include in the log file, tagging it as a debug message
        /// </summary>
        /// <remarks>The message is shown in dark gray in the console.</remarks>
        /// <param name="statusMessage">Status message</param>
        /// <param name="writeToLog">True to write to the log file; false to only display at console</param>
        protected static void LogDebug(string statusMessage, bool writeToLog = true)
        {
            LogTools.LogDebug(statusMessage, writeToLog);
        }

        /// <summary>
        /// Log an error message
        /// </summary>
        /// <param name="errorMessage">Error message</param>
        /// <param name="logToDb">When true, log the message to the database and the local log file</param>
        protected static void LogError(string errorMessage, bool logToDb = false)
        {
            LogTools.LogError(errorMessage, null, logToDb);
        }

        /// <summary>
        /// Log an error message and exception
        /// </summary>
        /// <param name="errorMessage">Error message (do not include ex.message)</param>
        /// <param name="ex">Exception to log</param>
        protected static void LogError(string errorMessage, Exception ex)
        {
            LogTools.LogError(errorMessage, ex);
        }

        /// <summary>
        /// Show a status message at the console and optionally include in the log file
        /// </summary>
        /// <param name="statusMessage">Status message</param>
        /// <param name="isError">True if this is an error</param>
        /// <param name="writeToLog">True to write to the log file; false to only display at console</param>
        public static void LogMessage(string statusMessage, bool isError = false, bool writeToLog = true)
        {
            if (writeToLog)
            {
                if (isError)
                    LogTools.LogError(statusMessage);
                else
                    LogTools.LogMessage(statusMessage);
            }
            else
            {
                if (isError)
                    ConsoleMsgUtils.ShowErrorCustom(statusMessage, false);
                else
                    Console.WriteLine(statusMessage);
            }
        }

        /// <summary>
        /// Log a warning message
        /// </summary>
        /// <param name="warningMessage">Warning message</param>
        /// <param name="logToDb">When true, log the message to the database and the local log file</param>
        protected static void LogWarning(string warningMessage, bool logToDb = false)
        {
            LogTools.LogWarning(warningMessage, logToDb);
        }
    }
}
