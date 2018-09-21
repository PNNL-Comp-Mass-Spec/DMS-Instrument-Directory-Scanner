//*********************************************************************************************************
// Written by Dave Clark for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Copyright 2009, Battelle Memorial Institute
// Created 07/27/2009
//
//*********************************************************************************************************

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using PRISM;
using PRISM.Logging;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Master processing class
    /// </summary>
    public class clsMainProcess : clsLoggerBase
    {
        #region "Constants"

        private const string DEFAULT_BASE_LOGFILE_NAME = @"Logs\InstDirScanner";

        private const string MGR_PARAM_DEFAULT_DMS_CONN_STRING = "MgrCnfgDbConnectStr";

        #endregion

        #region "Member variables"

        private readonly string m_MgrExeName;

        private readonly string m_MgrDirectoryPath;

        private clsMgrSettings m_MgrSettings;

        static clsStatusFile m_StatusFile;

        private clsMessageHandler m_MsgHandler;

        #endregion

        #region "Properties"

        /// <summary>
        /// When true, ignore Bionet instruments
        /// </summary>
        public bool NoBionet { get; set; }

        /// <summary>
        /// When true, preview the stats but don't change any instrument stat files
        /// </summary>
        public bool PreviewMode { get; set; }

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        public clsMainProcess()
        {
            var exeInfo = new FileInfo(GetAppPath());
            m_MgrExeName = exeInfo.Name;
            m_MgrDirectoryPath = exeInfo.DirectoryName;
        }

        /// <summary>
        /// Initializes the manager settings and classes
        /// </summary>
        /// <returns>True for success; False if error occurs</returns>
        /// <remarks></remarks>
        public bool InitMgr()
        {
            // Define the default logging info
            // This will get updated below
            LogTools.CreateFileLogger(DEFAULT_BASE_LOGFILE_NAME, BaseLogger.LogLevels.DEBUG);

            // Create a database logger connected to DMS5
            // Once the initial parameters have been successfully read,
            // we update the dbLogger to use the connection string read from the Manager Control DB
            string defaultDmsConnectionString;

            // Open DMS_InstDirScanner.exe.config to look for setting MgrCnfgDbConnectStr, so we know which server to log to by default
            var dmsConnectionStringFromConfig = GetXmlConfigDefaultConnectionString();

            if (string.IsNullOrWhiteSpace(dmsConnectionStringFromConfig))
            {
                // Use the hard-coded default that points to Gigasax
                defaultDmsConnectionString = Properties.Settings.Default.MgrCnfgDbConnectStr;
            }
            else
            {
                // Use the connection string from DMS_InstDirScanner.exe.config
                defaultDmsConnectionString = dmsConnectionStringFromConfig;
            }

            // ReSharper disable once ConditionIsAlwaysTrueOrFalse
            LogTools.CreateDbLogger(defaultDmsConnectionString, "InstDirScan: " + System.Net.Dns.GetHostName());

            // Get the manager settings
            try
            {
                m_MgrSettings = new clsMgrSettings();
            }
            catch
            {
                // Failures are logged by clsMgrSettings to local emergency log file
                return false;
            }

            // Setup the loggers
            var logFileNameBase = m_MgrSettings.GetParam("LogFileName", "InstDirScanner");

            BaseLogger.LogLevels logLevel;
            if (int.TryParse(m_MgrSettings.GetParam("DebugLevel"), out var debugLevel))
            {
                logLevel = (BaseLogger.LogLevels)debugLevel;
            }
            else
            {
                logLevel = BaseLogger.LogLevels.INFO;
            }

            LogTools.CreateFileLogger(logFileNameBase, logLevel);

            // Typically:
            // Data Source=gigasax;Initial Catalog=DMS5;Integrated Security=SSPI;
            var logCnStr = m_MgrSettings.GetParam("ConnectionString");
            var moduleName = m_MgrSettings.GetParam("ModuleName");
            LogTools.CreateDbLogger(logCnStr, moduleName);

            // Make the initial log entry
            var msg = "=== Started Instrument Directory Scanner V" + GetAppVersion() + " ===== ";

            LogTools.LogMessage(msg);

            // Setup the message queue

            m_MsgHandler = new clsMessageHandler
            {
                BrokerUri = m_MgrSettings.GetParam("MessageQueueURI"),
                StatusTopicName = m_MgrSettings.GetParam("MessageQueueTopicMgrStatus"),
                MgrSettings = m_MgrSettings
            };

            if (!m_MsgHandler.Init())
            {
                // Most error messages provided by .Init method, but debug message is here for program tracking
                LogDebug("Message handler init error");
                return false;
            }

            LogDebug("Message handler initialized");

            // Optional: Connect message handler events
            // m_MsgHandler.CommandReceived += OnMsgHandler_CommandReceived;
            // m_MsgHandler.BroadcastReceived += OnMsgHandler_BroadcastReceived;

            // Setup the status file class
            var appPath = GetAppPath();
            var fInfo = new FileInfo(appPath);

            string statusFileNameLoc;
            if (fInfo.DirectoryName == null)
                statusFileNameLoc = "Status.xml";
            else
                statusFileNameLoc = Path.Combine(fInfo.DirectoryName, "Status.xml");

            m_StatusFile = new clsStatusFile(statusFileNameLoc, m_MsgHandler)
            {
                LogToMsgQueue = m_MgrSettings.GetParam("LogStatusToMessageQueue", false),
                MgrName = m_MgrSettings.ManagerName,
                MgrStatus = clsStatusFile.EnumMgrStatus.Running
            };

            AttachEvents(m_StatusFile);

            m_StatusFile.WriteStatusFile();

            LogDebug("Status file init complete");

            // Everything worked
            return true;
        }

        /// <summary>
        /// Do a directory scan
        /// </summary>
        /// <remarks></remarks>
        public void DoDirectoryScan()
        {
            try
            {
                // Check to see if manager is active
                if (!bool.Parse(m_MgrSettings.GetParam("MgrActive")))
                {
                    var message = "Program disabled in manager control DB";
                    ConsoleMsgUtils.ShowWarning(message);
                    LogTools.LogMessage(message);
                    LogTools.LogMessage("===== Closing Inst Dir Scanner =====");
                    m_StatusFile.UpdateDisabled(false);
                    return;
                }

                if (!bool.Parse(m_MgrSettings.GetParam("MgrActive_local")))
                {
                    var message = "Program disabled locally";
                    ConsoleMsgUtils.ShowWarning(message);
                    LogTools.LogMessage(message);
                    LogTools.LogMessage("===== Closing Inst Dir Scanner =====");
                    m_StatusFile.UpdateDisabled(true);
                    return;
                }

                var workDir = m_MgrSettings.GetParam("WorkDir", string.Empty);
                if (string.IsNullOrWhiteSpace(workDir))
                {
                    LogFatalError("Manager parameter \'WorkDir\' is not defined");
                    return;
                }

                // Verify output directory can be found
                if (!Directory.Exists(workDir))
                {
                    LogFatalError("Output directory not found: " + workDir);
                    return;
                }

                // Get list of instruments from DMS
                var instList = GetInstrumentList();
                if (instList == null)
                {
                    LogFatalError("No instruments");
                    return;
                }

                // Scan the directories
                var scanner = new clsDirectoryTools(NoBionet, PreviewMode);
                AttachEvents(scanner);

                scanner.PerformDirectoryScans(instList, workDir, m_MgrSettings, m_StatusFile);

                // All finished, so clean up and exit
                LogMessage("Scanning complete");
                LogTools.LogMessage("===== Closing Inst Dir Scanner =====");
                m_StatusFile.UpdateStopped(false);
            }
            catch (Exception ex)
            {
                LogError("Error in DoDirectoryScan", ex);
            }

        }

        /// <summary>
        /// Returns the full path to the executing .Exe or .Dll
        /// </summary>
        /// <returns>File path</returns>
        /// <remarks></remarks>
        public static string GetAppPath()
        {
            return PRISM.FileProcessor.ProcessFilesOrFoldersBase.GetAppPath();
        }

        /// <summary>
        /// Returns the .NET assembly version followed by the program date
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string GetAppVersion()
        {
            return PRISM.FileProcessor.ProcessFilesOrFoldersBase.GetEntryOrExecutingAssembly().GetName().Version.ToString();
        }

        private List<clsInstData> GetInstrumentList()
        {
            LogMessage("Getting instrument list");
            var columns = new List<string> {
                "vol",
                "path",
                "method",
                "Instrument" };

            var sqlQuery = "SELECT " + string.Join(",", columns) + " FROM V_Instrument_Source_Paths";

            var connectionString = m_MgrSettings.GetParam("ConnectionString");
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                LogError("Connection string is empty; cannot retrieve manager parameters");
                return null;
            }

            var dbTools = new DBTools(connectionString);
            AttachEvents(dbTools);

            // Get a table containing the active instruments
            var success = dbTools.GetQueryResults(sqlQuery, out var lstResults, "GetInstrumentList");

            // Verify valid data found
            if (!success || lstResults == null)
            {
                LogError("Unable to retrieve instrument list");
                return null;
            }

            if (lstResults.Count < 1)
            {
                LogError("No instruments found in V_Instrument_Source_Paths using " + connectionString);
                return null;
            }

            var colMapping = dbTools.GetColumnMapping(columns);

            // Create a list of all instrument data
            var instrumentList = new List<clsInstData>();
            try
            {
                foreach (var result in lstResults)
                {
                    var instrumentInfo = new clsInstData
                    {
                        CaptureMethod = dbTools.GetColumnValue(result, colMapping, "method"),
                        InstName = dbTools.GetColumnValue(result, colMapping, "Instrument"),
                        StoragePath = dbTools.GetColumnValue(result, colMapping, "path"),
                        StorageVolume = dbTools.GetColumnValue(result, colMapping, "vol")
                    };

                    instrumentList.Add(instrumentInfo);
                }

                LogTools.LogDebug("Retrieved instrument list");
                return instrumentList;
            }
            catch (Exception ex)
            {
                LogError("Exception filling instrument list", ex);
                return null;
            }
        }


        /// <summary>
        /// Extract the value MgrCnfgDbConnectStr from DMS_InstDirScanner.exe.config
        /// </summary>
        /// <returns></returns>
        private string GetXmlConfigDefaultConnectionString()
        {
            return GetXmlConfigFileSetting(MGR_PARAM_DEFAULT_DMS_CONN_STRING);
        }

        /// <summary>
        /// Extract the value for the given setting from DMS_InstDirScanner.exe.config
        /// </summary>
        /// <returns>Setting value if found, otherwise an empty string</returns>
        /// <remarks>Uses a simple text reader in case the file has malformed XML</remarks>
        private string GetXmlConfigFileSetting(string settingName)
        {

            if (string.IsNullOrWhiteSpace(settingName))
                throw new ArgumentException("Setting name cannot be blank", nameof(settingName));

            try
            {
                var configFilePath = Path.Combine(m_MgrDirectoryPath, m_MgrExeName + ".config");
                var configFile = new FileInfo(configFilePath);

                if (!configFile.Exists)
                {
                    LogError("File not found: " + configFilePath);
                    return string.Empty;
                }

                var configXml = new StringBuilder();

                // Open DMS_InstDirScanner.exe.config using a simple text reader in case the file has malformed XML

                using (var reader = new StreamReader(new FileStream(configFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    while (!reader.EndOfStream)
                    {
                        var dataLine = reader.ReadLine();
                        if (string.IsNullOrWhiteSpace(dataLine))
                            continue;

                        configXml.Append(dataLine);
                    }
                }

                var matcher = new Regex(settingName + ".+?<value>(?<ConnString>.+?)</value>", RegexOptions.IgnoreCase);

                var match = matcher.Match(configXml.ToString());

                if (match.Success)
                    return match.Groups["ConnString"].Value;

                LogError(settingName + " setting not found in " + configFilePath);
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogError("Exception reading setting " + settingName + " in DMS_InstDirScanner.exe.config", ex);
                return string.Empty;
            }

        }

        private void LogFatalError(string errorMessage)
        {
            LogError(errorMessage);
            LogTools.LogMessage("===== Closing Inst Dir Scanner =====");
            m_StatusFile.UpdateStopped(true);

        }

        #region "Event Handlers"

        private void AttachEvents(EventNotifier objClass)
        {
            objClass.ErrorEvent += ErrorEventHandler;
            objClass.StatusEvent += MessageEventHandler;
            objClass.WarningEvent += WarningEventHandler;
        }

        private void ErrorEventHandler(string message, Exception ex)
        {
            LogError(message);
        }

        private void MessageEventHandler(string message)
        {
            LogMessage(message);
        }

        private void WarningEventHandler(string message)
        {
            LogWarning(message);
        }

        #endregion
    }


}