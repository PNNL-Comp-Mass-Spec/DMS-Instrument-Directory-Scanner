using System;
using PRISM;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class that starts application execution
    /// </summary>
    static class Program
    {

        static clsMainProcess m_MainProcess;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {

            // Start the main program running
            try
            {
                m_MainProcess = new clsMainProcess();
                if (!m_MainProcess.InitMgr())
                {
                    PRISM.Logging.FileLogger.FlushPendingMessages();
                    clsProgRunner.SleepMilliseconds(1500);
                    return;
                }
                m_MainProcess.DoDirectoryScan();
            }
            catch (Exception ex)
            {
                var errMsg = "Critical exception starting application: " + ex.Message;
                ConsoleMsgUtils.ShowWarning(errMsg + "; " + clsStackTraceFormatter.GetExceptionStackTrace(ex));
                ConsoleMsgUtils.ShowWarning("Exiting clsMainProcess.Main with error code = 1");
            }

            PRISM.Logging.FileLogger.FlushPendingMessages();

        }

    }
}
