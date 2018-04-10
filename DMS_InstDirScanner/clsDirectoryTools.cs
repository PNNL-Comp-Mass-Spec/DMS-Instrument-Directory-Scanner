//*********************************************************************************************************
// Written by Dave Clark for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Copyright 2009, Battelle Memorial Institute
// Created 01/01/2009
//
//*********************************************************************************************************

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PRISM;
using PRISM.Logging;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Handles all directory access tasks
    /// </summary>
    internal class clsDirectoryTools : clsEventNotifier
    {

        #region "Member variables"

        int mDebugLevel = 1;
        string mMostRecentIOErrorInstrument;
        private readonly clsFileTools mFileTools;

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        /// <remarks></remarks>
        public clsDirectoryTools()
        {
            mMostRecentIOErrorInstrument = string.Empty;
            mFileTools = new clsFileTools();
        }

        public bool PerformDirectoryScans(List<clsInstData> instList, string outDirectoryPath, clsMgrSettings mgrSettings, clsStatusFile progStatus)
        {
            var instCounter = 0;
            var instCount = instList.Count;

            mDebugLevel = mgrSettings.GetParam("debuglevel", 1);

            progStatus.TaskStartTime = DateTime.UtcNow;

            foreach (var instrument in instList)
            {
                try
                {
                    instCounter++;
                    var progress = 100 * instCounter / (float)instCount;
                    progStatus.UpdateAndWrite(progress);

                    OnStatusEvent("Scanning directory for instrument " + instrument.InstName);

                    var swOutFile = CreateOutputFile(instrument.InstName, outDirectoryPath, out var fiSourceFile);
                    if (swOutFile == null)
                    {
                        return false;
                    }

                    var directoryExists = GetDirectoryData(instrument, swOutFile, mgrSettings);

                    swOutFile.Close();

                    if (directoryExists)
                    {
                        // Copy the file to the MostRecentValid directory
                        try
                        {
                            var diTargetDirectory = new DirectoryInfo(Path.Combine(outDirectoryPath, "MostRecentValid"));

                            if (!diTargetDirectory.Exists)
                            {
                                diTargetDirectory.Create();
                            }

                            fiSourceFile.CopyTo(Path.Combine(diTargetDirectory.FullName, fiSourceFile.Name), true);
                        }
                        catch (Exception ex)
                        {
                            OnErrorEvent("Exception copying to MostRecentValid directory", ex);
                        }

                    }

                }
                catch (Exception ex)
                {
                    LogCriticalError("Error finding files for " + instrument.InstName + " in PerformDirectoryScans: " + ex.Message);
                }

            }

            return true;
        }

        private StreamWriter CreateOutputFile(string instName, string outFileDir, out FileInfo fiStatusFile)
        {
            fiStatusFile = new FileInfo(Path.Combine(outFileDir, instName + "_source.txt"));

            // Make a backup copy of the existing file
            if (fiStatusFile.Exists && fiStatusFile.Directory != null)
            {
                try
                {
                    var backupDirectory = new DirectoryInfo(Path.Combine(fiStatusFile.Directory.FullName, "PreviousCopy"));
                    if (!backupDirectory.Exists)
                    {
                        backupDirectory.Create();
                    }

                    fiStatusFile.CopyTo(Path.Combine(backupDirectory.FullName, fiStatusFile.Name), true);
                }
                catch (Exception ex)
                {
                    OnErrorEvent("Exception copying " + fiStatusFile.Name + "to PreviousCopy directory", ex);
                }

                if (!mFileTools.DeleteFileWithRetry(fiStatusFile, out var backupErrorMessage))
                {
                    LogErrorToDatabase(backupErrorMessage);
                    return null;
                }

            }

            // Create the new file; try up to 3 times
            var retriesRemaining = 3;
            var errorMessage = string.Empty;

            while (retriesRemaining >= 0)
            {
                retriesRemaining--;
                try
                {
                    var swOutFile = new StreamWriter(new FileStream(fiStatusFile.FullName, FileMode.Create, FileAccess.Write, FileShare.Read));

                    // The file always starts with a blank line
                    swOutFile.WriteLine();
                    return swOutFile;
                }
                catch (Exception ex)
                {
                    errorMessage = ex.Message;
                    // Delay for 1 second before trying again
                    clsProgRunner.SleepMilliseconds(1000);
                }

            }

            OnErrorEvent("Exception creating output file " + fiStatusFile.FullName + ": " + errorMessage);
            return null;
        }

        /// <summary>
        /// Query the files and directories on the instrument's shared data directory
        /// </summary>
        /// <param name="intrumentData"></param>
        /// <param name="swOutFile"></param>
        /// <param name="mgrSettings"></param>
        /// <returns>True on success, false if the target directory is not found</returns>
        /// <remarks></remarks>
        private bool GetDirectoryData(clsInstData intrumentData, TextWriter swOutFile, clsMgrSettings mgrSettings)
        {
            var connected = false;
            var remoteDirectoryPath = Path.Combine(intrumentData.StorageVolume, intrumentData.StoragePath);
            ShareConnector shareConn = null;

            string userDescription;

            // If this is a machine on bionet, set up a connection
            if (intrumentData.CaptureMethod.ToLower() == "secfso")
            {
                // Typically user ftms (not LCMSOperator)
                var bionetUser = mgrSettings.GetParam("bionetuser");

                if (!bionetUser.Contains("\\"))
                {
                    // Prepend this computer's name to the username
                    bionetUser = Environment.MachineName + "\\" + bionetUser;
                }

                shareConn = new ShareConnector(remoteDirectoryPath, bionetUser, DecodePassword(mgrSettings.GetParam("bionetpwd")));
                connected = shareConn.Connect();

                userDescription = " as user " + bionetUser;
                if (!connected)
                {
                    OnErrorEvent("Could not connect to " + remoteDirectoryPath + userDescription + "; error code " + shareConn.ErrorMessage);
                }
                else if (mDebugLevel >= 5)
                {
                    OnStatusEvent(" ... connected to " + remoteDirectoryPath + userDescription);
                }

            }
            else
            {
                userDescription = " as user " + Environment.UserName;
                if (remoteDirectoryPath.ToLower().Contains(".bionet"))
                {
                    OnWarningEvent("Warning: Connection to a bionet share should probably use \'secfso\'; " +
                                    "currently configured to use \'fso\' for " + remoteDirectoryPath);
                }

            }

            var instrumentDataDirectory = new DirectoryInfo(remoteDirectoryPath);

            OnStatusEvent("Reading " + intrumentData.InstName + ", Directory " + remoteDirectoryPath + userDescription);

            // List the directory path and current date/time on the first line
            // Will look like this:
            // (Directory: \\VOrbiETD04.bionet\ProteomicsData\ at 2012-01-23 2:15 PM)
            WriteToOutput(swOutFile, "Directory: " + remoteDirectoryPath + " at " + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt"));

            var directoryExists = Directory.Exists(remoteDirectoryPath);
            if (!directoryExists)
            {
                OnStatusEvent("Path not found: " + remoteDirectoryPath);
                WriteToOutput(swOutFile, "(Directory does not exist)");
            }
            else
            {
                var directories = instrumentDataDirectory.GetDirectories().ToList();
                var files = instrumentDataDirectory.GetFiles().ToList();
                var archiveCount = 0;

                foreach (var datasetDirectory in directories)
                {
                    var totalSizeBytes = GetDirectorySize(intrumentData.InstName, datasetDirectory);
                    if (datasetDirectory.Name.StartsWith("x_"))
                        archiveCount++;

                    WriteToOutput(swOutFile, "Dir ", datasetDirectory.Name, FileSizeToText(totalSizeBytes));
                }

                foreach (var datasetFile in files)
                {
                    var fileSizeText = FileSizeToText(datasetFile.Length);
                    if (datasetFile.Name.StartsWith("x_"))
                        archiveCount++;

                    WriteToOutput(swOutFile, "File ", datasetFile.Name, fileSizeText);
                }

                var fileCountText = files.Count == 1 ? "file" : "files";
                var dirCountText = directories.Count == 1 ? "directory" : "directories";
                var archiveCountText = archiveCount == 1 ? "has" : "have";

                // Construct the file stats message, for example:
                // Found 2 directories and 12 files; 8 have been archived

                var fileStats = string.Format("Found {0} {1} and {2} {3}", files.Count, fileCountText, directories.Count, dirCountText);
                if (archiveCount > 0)
                {
                    fileStats += string.Format("; {0} {1} been archived", archiveCount, archiveCountText);
                }

                OnStatusEvent(fileStats);

            }

            // If this was a bionet machine, disconnect
            if (connected)
            {
                if (!shareConn.Disconnect())
                {
                    OnErrorEvent("Could not disconnect from " + remoteDirectoryPath);
                }

            }

            return directoryExists;
        }

        private string FileSizeToText(long fileSizeBytes)
        {
            var fileSize = (double)fileSizeBytes;
            var fileSizeIterator = 0;
            while (fileSize > 1024 && fileSizeIterator < 3)
            {
                fileSize /= 1024;
                fileSizeIterator++;
            }

            string formatString;
            if (fileSize < 10)
            {
                formatString = "0.0";
            }
            else
            {
                formatString = "0";
            }

            var fileSizeText = fileSize.ToString(formatString);
            string fileSizeUnits;

            switch (fileSizeIterator)
            {
                case 0:
                    fileSizeUnits = "bytes";
                    break;
                case 1:
                    fileSizeUnits = "KB";
                    break;
                case 2:
                    fileSizeUnits = "MB";
                    break;
                case 3:
                    fileSizeUnits = "GB";
                    break;
                default:
                    fileSizeUnits = "???";
                    break;
            }
            return fileSizeText + " " + fileSizeUnits;
        }

        private long GetDirectorySize(string instrumentName, DirectoryInfo datasetDirectory)
        {
            long totalSizeBytes = 0;
            try
            {
                var files = datasetDirectory.GetFiles("*").ToList();
                foreach (var currentFile in files)
                {
                    totalSizeBytes = (totalSizeBytes + currentFile.Length);
                }

                // Do not use SearchOption.AllDirectories in case there are security issues with one or more of the subdirectories
                var subDirs = datasetDirectory.GetDirectories("*").ToList();
                foreach (var subDirectory in subDirs)
                {
                    // Recursively call this function
                    var subDirSizeBytes = GetDirectorySize(instrumentName, subDirectory);
                    totalSizeBytes += subDirSizeBytes;
                }

            }
            catch (UnauthorizedAccessException)
            {
                // Ignore this error
            }
            catch (IOException)
            {
                LogIOError(instrumentName,
                           "IOException determining directory size for " + instrumentName + ", " +
                           "directory " + datasetDirectory.Name + "; ignoring additional errors for this instrument");
            }
            catch (Exception ex)
            {
                LogCriticalError("Error determining directory size for " + instrumentName + ", " +
                                 "directory " + datasetDirectory.Name + ": " + ex.Message);
            }

            return totalSizeBytes;
        }

        private void WriteToOutput(TextWriter swOutFile, string field1, string field2 = "", string field3 = "" )
        {
            var dataLine = field1 + '\t' + field2 + '\t' + field3;
            swOutFile.WriteLine(dataLine);
        }

        private string DecodePassword(string encodedPwd)
        {
            return Pacifica.Core.Utilities.DecodePassword(encodedPwd);
        }

        /// <summary>
        /// Logs an error message to the local log file, unless it is currently between midnight and 1 am
        /// </summary>
        /// <param name="errorMessage"></param>
        /// <remarks></remarks>
        private void LogCriticalError(string errorMessage)
        {
            if (DateTime.Now.Hour == 0)
            {
                // Log this error to the database if it is between 12 am and 1 am
                LogErrorToDatabase(errorMessage);
            }
            else
            {
                OnErrorEvent(errorMessage);
            }

        }

        private void LogErrorToDatabase(string errorMessage)
        {
            LogTools.LogError(errorMessage, null, true);
            OnErrorEvent(errorMessage);
        }

        private void LogIOError(string instrumentName, string errorMessage)
        {
            if (!string.IsNullOrEmpty(mMostRecentIOErrorInstrument) && mMostRecentIOErrorInstrument == instrumentName)
            {
                return;
            }

            mMostRecentIOErrorInstrument = instrumentName;
            OnWarningEvent(errorMessage);
        }

    }

}