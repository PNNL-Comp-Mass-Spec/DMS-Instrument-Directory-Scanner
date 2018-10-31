using System;
using System.Collections.Generic;
using System.Linq;
using PRISM;
using PRISM.FileProcessor;
using PRISM.Logging;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class that starts application execution
    /// </summary>
    internal class Program
    {
        public const string PROGRAM_DATE = "October 30, 2018";

        private static bool mNoBionet;

        private static bool mPreviewMode;

        /// <summary>
        /// Entry method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main(string[] args)
        {
            var commandLineParser = new clsParseCommandLine();

            mNoBionet = false;

            mPreviewMode = false;

            try
            {
                bool validArgs;

                // Parse the command line options
                if (commandLineParser.ParseCommandLine())
                {
                    validArgs = SetOptionsUsingCommandLineParameters(commandLineParser);
                }
                else if (commandLineParser.NoParameters)
                {
                    validArgs = true;
                }
                else
                {
                    if (commandLineParser.NeedToShowHelp)
                    {
                        ShowProgramHelp();
                    }
                    else
                    {
                        ConsoleMsgUtils.ShowWarning("Error parsing the command line arguments");
                        clsParseCommandLine.PauseAtConsole(750);
                    }

                    return -1;
                }

                if (commandLineParser.NeedToShowHelp || !validArgs)
                {
                    ShowProgramHelp();
                    return -1;
                }

                var mainProcess = new MainProcess
                {
                    NoBionet = mNoBionet,
                    PreviewMode = mPreviewMode
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

        private static bool SetOptionsUsingCommandLineParameters(clsParseCommandLine commandLineParser)
        {
            // Returns True if no problems; otherwise, returns false
            var validParameters = new List<string>
            {
                "NoBionet",
                "Preview"
            };

            try
            {
                // Make sure no invalid parameters are present
                if (commandLineParser.InvalidParametersPresent(validParameters))
                {
                    ShowErrorMessage("Invalid command line parameters",
                        (from item in commandLineParser.InvalidParameters(validParameters) select "/" + item).ToList());

                    return false;
                }

                // Query commandLineParser to see if various parameters are present
                if (commandLineParser.IsParameterPresent("NoBionet"))
                {
                    mNoBionet = true;
                }

                if (commandLineParser.IsParameterPresent("Preview"))
                {
                    mPreviewMode = true;
                }

                return true;

            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error parsing the command line parameters: ", ex);
            }

            return false;
        }

        private static void ShowErrorMessage(string message, Exception ex)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }

        private static void ShowErrorMessage(string title, IEnumerable<string> errorMessages)
        {
            ConsoleMsgUtils.ShowErrors(title, errorMessages);
        }

        private static void ShowProgramHelp()
        {
            try
            {
                var exePath = MainProcess.GetAppPath();

                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "This program finds the files and directories in the source folder for active DMS instruments. " +
                    "It creates a text file for each instrument on a central share, listing the files and directories." +
                    "On the DMS website, the helper_inst_source page file reads these text files show DMS users dataset" +
                    "files and directories on the instruments, for example https://dms2.pnl.gov/helper_inst_source/view/QExactP02"));
                Console.WriteLine();
                Console.WriteLine("Program syntax:");
                Console.WriteLine(exePath + " [/NoBionet] [/Preview]");
                Console.WriteLine();
                Console.WriteLine("Use /NoBionet to skip instruments on bionet");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /Preview to search for files and directories, but not update any files on the DMS_InstSourceDirScans share"));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Program written by Dave Clark and Matthew Monroe for the Department of Energy (PNNL, Richland, WA)"));
                Console.WriteLine();
                Console.WriteLine("Version: " + ProcessFilesOrDirectoriesBase.GetAppVersion(PROGRAM_DATE));
                Console.WriteLine();
                Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov");
                Console.WriteLine("Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Licensed under the 2-Clause BSD License; you may not use this file except in compliance with the License. " +
                    "You may obtain a copy of the License at https://opensource.org/licenses/BSD-2-Clause"));
                Console.WriteLine();

                // Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
                clsParseCommandLine.PauseAtConsole(750);
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error displaying the program syntax", ex);
            }

        }

    }
}
