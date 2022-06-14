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
using PRISM.AppSettings;
using PRISM.Logging;
using PRISMDatabaseUtils;
using PRISMDatabaseUtils.AppSettings;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Master processing class
    /// </summary>
    public class MainProcess : LoggerBase
    {
        // Ignore Spelling: Bionet

        private const string DEFAULT_BASE_LOGFILE_NAME = @"Logs\InstDirScanner";

        private const string MGR_PARAM_DEFAULT_DMS_CONN_STRING = "MgrCnfgDbConnectStr";

        private readonly string m_MgrExeName;

        private readonly string m_MgrDirectoryPath;

        private MgrSettingsDB m_MgrSettings;

        private static StatusFile m_StatusFile;

        private MessageHandler m_MsgHandler;

        /// <summary>
        /// When true, ignore Bionet instruments
        /// </summary>
        public bool NoBionet { get; set; }

        /// <summary>
        /// When true, preview the stats but don't change any instrument stat files
        /// </summary>
        public bool PreviewMode { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public MainProcess()
        {
            var exeInfo = new FileInfo(GetAppPath());
            m_MgrExeName = exeInfo.Name;
            m_MgrDirectoryPath = exeInfo.DirectoryName;
        }

        /// <summary>
        /// Initializes the manager settings and classes
        /// </summary>
        /// <returns>True for success; False if error occurs</returns>
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

            var hostName = System.Net.Dns.GetHostName();
            var applicationName = "InstDirScanner_" + hostName;
            var defaultDbLoggerConnectionString = DbToolsFactory.AddApplicationNameToConnectionString(defaultDmsConnectionString, applicationName);

            // ReSharper disable once ConditionIsAlwaysTrueOrFalse
            LogTools.CreateDbLogger(defaultDbLoggerConnectionString, "InstDirScan: " + hostName);

            // Get the manager settings
            try
            {
                var defaultSettings = new Dictionary<string, string>
                {
                    {MgrSettings.MGR_PARAM_MGR_CFG_DB_CONN_STRING, Properties.Settings.Default.MgrCnfgDbConnectStr},
                    {MgrSettings.MGR_PARAM_MGR_ACTIVE_LOCAL, Properties.Settings.Default.MgrActive_Local.ToString()},
                    {MgrSettings.MGR_PARAM_MGR_NAME, Properties.Settings.Default.MgrName},
                    {MgrSettings.MGR_PARAM_USING_DEFAULTS, Properties.Settings.Default.UsingDefaults.ToString()}
                };

                m_MgrSettings = new MgrSettingsDB();
                RegisterEvents(m_MgrSettings);
                m_MgrSettings.CriticalErrorEvent += CriticalErrorEvent;

                var mgrExePath = PRISM.FileProcessor.ProcessFilesOrDirectoriesBase.GetAppPath();
                var localSettings = m_MgrSettings.LoadMgrSettingsFromFile(mgrExePath + ".config");

                if (localSettings == null)
                {
                    localSettings = defaultSettings;
                }
                else
                {
                    // Make sure the default settings exist and have valid values
                    foreach (var setting in defaultSettings)
                    {
                        if (!localSettings.TryGetValue(setting.Key, out var existingValue) ||
                            string.IsNullOrWhiteSpace(existingValue))
                        {
                            localSettings[setting.Key] = setting.Value;
                        }
                    }
                }

                Console.WriteLine();
                m_MgrSettings.ValidatePgPass(localSettings);

                var success = m_MgrSettings.LoadSettings(localSettings, true);
                if (!success)
                    return false;
            }
            catch
            {
                // Failures should have already been logged (or shown at the console)
                return false;
            }

            if (Environment.MachineName.StartsWith("monroe", StringComparison.OrdinalIgnoreCase))
            {
                var mgrPerspective = m_MgrSettings.GetParam("perspective");

                if (mgrPerspective.Equals("server", StringComparison.OrdinalIgnoreCase))
                {
                    m_MgrSettings.SetParam("perspective", "client");
                    Console.WriteLine("StoreParameters: Overriding manager perspective to be 'client' " +
                                      "because impersonating a server-based manager from an office computer");
                }
            }

            // Setup the loggers
            var logFileNameBase = m_MgrSettings.GetParam("LogFileName", "InstDirScanner");

            var debugLevel = m_MgrSettings.GetParam("DebugLevel", (int)BaseLogger.LogLevels.INFO);

            var logLevel = (BaseLogger.LogLevels)debugLevel;

            LogTools.CreateFileLogger(logFileNameBase, logLevel);

            // Typically:
            // Data Source=gigasax;Initial Catalog=DMS5;Integrated Security=SSPI;
            var logCnStr = m_MgrSettings.GetParam("ConnectionString");
            var moduleName = m_MgrSettings.GetParam("ModuleName");

            var dbLoggerConnectionString = DbToolsFactory.AddApplicationNameToConnectionString(logCnStr, m_MgrSettings.ManagerName);

            LogTools.CreateDbLogger(dbLoggerConnectionString, moduleName);

            // Make the initial log entry
            var msg = "=== Started Instrument Directory Scanner V" + GetAppVersion() + " === ";

            LogTools.LogMessage(msg);

            // Setup the message queue

            m_MsgHandler = new MessageHandler
            {
                BrokerUri = m_MgrSettings.GetParam("MessageQueueURI"),
                StatusTopicName = m_MgrSettings.GetParam("MessageQueueTopicMgrStatus"),
                MgrSettings = m_MgrSettings
            };

            if (!m_MsgHandler.Init())
            {
                // Most error messages provided by .Init method, but debug message is here for program tracking
                LogDebug("Message handler init error");
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

            m_StatusFile = new StatusFile(statusFileNameLoc, m_MsgHandler)
            {
                LogToMsgQueue = m_MgrSettings.GetParam("LogStatusToMessageQueue", false),
                MgrName = m_MgrSettings.ManagerName,
                MgrStatus = StatusFile.EnumMgrStatus.Running
            };

            RegisterEvents(m_StatusFile);

            m_StatusFile.WriteStatusFile();

            LogDebug("Status file init complete");

            // Everything worked
            return true;
        }

        /// <summary>
        /// Do a directory scan
        /// </summary>
        public void DoDirectoryScan()
        {
            try
            {
                // Check to see if manager is active
                if (!m_MgrSettings.GetParam("MgrActive", true))
                {
                    const string message = "Program disabled in manager control DB";
                    ConsoleMsgUtils.ShowWarning(message);
                    LogTools.LogMessage(message);
                    LogTools.LogMessage("===== Closing Inst Dir Scanner =====");
                    m_StatusFile.UpdateDisabled(false);
                    return;
                }

                if (!m_MgrSettings.GetParam(MgrSettings.MGR_PARAM_MGR_ACTIVE_LOCAL, true))
                {
                    const string message = "Program disabled locally";
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
                var scanner = new DirectoryTools(NoBionet, PreviewMode);
                RegisterEvents(scanner);

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
        public static string GetAppPath()
        {
            return PRISM.FileProcessor.ProcessFilesOrDirectoriesBase.GetAppPath();
        }

        /// <summary>
        /// Returns the .NET assembly version followed by the program date
        /// </summary>
        public static string GetAppVersion()
        {
            return PRISM.FileProcessor.ProcessFilesOrDirectoriesBase.GetEntryOrExecutingAssembly().GetName().Version.ToString();
        }

        private List<InstrumentData> GetInstrumentList()
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

            var connectionStringToUse = DbToolsFactory.AddApplicationNameToConnectionString(connectionString, m_MgrSettings.ManagerName);

            var dbTools = DbToolsFactory.GetDBTools(connectionStringToUse);
            RegisterEvents(dbTools);

            // Get a table containing the active instruments
            var success = dbTools.GetQueryResults(sqlQuery, out var lstResults);

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
            var instrumentList = new List<InstrumentData>();
            try
            {
                foreach (var result in lstResults)
                {
                    var instrumentInfo = new InstrumentData
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

        private void RegisterEvents(IEventNotifier sourceClass)
        {
            sourceClass.DebugEvent += DebugEventHandler;
            sourceClass.ErrorEvent += ErrorEventHandler;
            sourceClass.StatusEvent += MessageEventHandler;
            sourceClass.WarningEvent += WarningEventHandler;
        }

        private void CriticalErrorEvent(string message, Exception ex)
        {
            LogError(message, true);
        }

        private void DebugEventHandler(string message)
        {
            ConsoleMsgUtils.ShowDebugCustom(message, emptyLinesBeforeMessage: 0);
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
    }
}