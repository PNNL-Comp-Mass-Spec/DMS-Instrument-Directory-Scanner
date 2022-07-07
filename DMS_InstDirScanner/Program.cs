using System;
using System.Threading;
using PRISM;
using PRISM.FileProcessor;
using PRISM.Logging;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class that starts application execution
    /// </summary>
    internal static class Program
    {
        public const string PROGRAM_DATE = "July 7, 2022";

        /// <summary>
        /// Entry method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main(string[] args)
        {
            try
            {
                var exeName = System.IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);

                var parser = new CommandLineParser<CommandLineOptions>(exeName,
                    ProcessFilesOrDirectoriesBase.GetAppVersion(PROGRAM_DATE))
                {
                    ProgramInfo = "This program finds the files and directories in the source folder for active DMS instruments. " +
                                  "It creates a text file for each instrument on a central share, listing the files and directories." +
                                  "On the DMS website, the helper_inst_source page file reads these text files show DMS users dataset" +
                                  "files and directories on the instruments, for example https://dms2.pnl.gov/helper_inst_source/view/QExactP02",
                    ContactInfo =
                        "Program written by Dave Clark and Matthew Monroe for the Department of Energy (PNNL, Richland, WA)" + Environment.NewLine +
                        "E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov" + Environment.NewLine +
                        "Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics" + Environment.NewLine + Environment.NewLine +
                        "Licensed under the 2-Clause BSD License; you may not use this file except in compliance with the License. " +
                        "You may obtain a copy of the License at https://opensource.org/licenses/BSD-2-Clause"
                };

                var result = parser.ParseArgs(args, false);
                var options = result.ParsedResults;

                if (args.Length > 0 && !result.Success)
                {
                    if (parser.CreateParamFileProvided)
                    {
                        return 0;
                    }

                    // Delay for 1500 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
                    Thread.Sleep(1500);
                    return -1;
                }

                var mainProcess = new MainProcess
                {
                    NoBionet = options.NoBionet,
                    PreviewMode = options.PreviewMode
                };

                try
                {
                    if (!mainProcess.InitMgr())
                    {
                        FileLogger.FlushPendingMessages();
                        ProgRunner.SleepMilliseconds(1500);
                        return -2;
                    }
                }
                catch (Exception ex)
                {
                    ShowErrorMessage("Exception thrown by InitMgr: ", ex);

                    LogTools.FlushPendingMessages();
                    return -1;
                }

                mainProcess.DoDirectoryScan();
                LogTools.FlushPendingMessages();
                return 0;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error occurred in modMain->Main", ex);
                LogTools.FlushPendingMessages();
                return -1;
            }
        }

        private static void ShowErrorMessage(string message, Exception ex)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }
    }
}
