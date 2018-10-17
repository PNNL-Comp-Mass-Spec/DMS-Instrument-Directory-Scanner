
//*********************************************************************************************************
// Written by Dave Clark for the US Department of Energy
// Pacific Northwest National Laboratory, Richland, WA
// Copyright 2009, Battelle Memorial Institute
// Created 06/16/2009
//*********************************************************************************************************

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using PRISM;

namespace DMS_InstDirScanner
{
    /// <summary>
    /// Class for loading, storing and accessing manager parameters.
    /// </summary>
    /// <remarks>
    /// Loads initial settings from local config file, then checks to see if remainder of settings should be
    /// loaded or manager set to inactive. If manager active, retrieves remainder of settings from manager
    /// parameters database.
    /// </remarks>
    public class MgrSettings : EventNotifier
    {
        #region "Constants"

        /// <summary>
        /// Status message for when the manager is deactivated locally
        /// </summary>
        /// <remarks>Used when MgrActive_Local is False in AppName.exe.config</remarks>
        public const string DEACTIVATED_LOCALLY = "Manager deactivated locally";

        /// <summary>
        /// Manager parameter: config database connection string
        /// </summary>
        public const string MGR_PARAM_MGR_CFG_DB_CONN_STRING = "MgrCnfgDbConnectStr";
        /// <summary>
        /// Manager parameter: manager active
        /// </summary>
        /// <remarks>Defined in AppName.exe.config</remarks>
        public const string MGR_PARAM_MGR_ACTIVE_LOCAL = "MgrActive_Local";

        /// <summary>
        /// Manager parameter: manager name
        /// </summary>
        public const string MGR_PARAM_MGR_NAME = "MgrName";

        /// <summary>
        /// Manager parameter: using defaults flag
        /// </summary>
        public const string MGR_PARAM_USING_DEFAULTS = "UsingDefaults";

        #endregion

        #region "Properties"

        /// <summary>
        /// Error message
        /// </summary>
        public string ErrMsg { get; private set; } = "";

        /// <summary>
        /// Manager name
        /// </summary>
        public string ManagerName => GetParam(MGR_PARAM_MGR_NAME, Environment.MachineName + "_Undefined-Manager");

        /// <summary>
        /// This will be true after the parameter have been successfully loaded from the database
        /// </summary>
        public bool ParamsLoadedFromDB { get; private set; }

        /// <summary>
        /// Dictionary of manager parameters
        /// </summary>
        public Dictionary<string, string> MgrParams { get; }

        #endregion

        #region "Events"

        /// <summary>
        /// Important error event (only raised if ParamsLoadedFromDB is false)
        /// </summary>
        public ErrorEventEventHandler CriticalErrorEvent;

        #endregion

        #region "Methods"

        /// <summary>
        /// Constructor
        /// </summary>
        public MgrSettings()
        {
            if (!LoadSettings())
            {
                if (string.Equals(ErrMsg, DEACTIVATED_LOCALLY))
                    throw new ApplicationException(DEACTIVATED_LOCALLY);
            ParamsLoadedFromDB = false;
            MgrParams = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                throw new ApplicationException("Unable to initialize manager settings class: " + ErrMsg);
            }
        }

        /// <summary>
        /// Load manager settings from the config file
        /// </summary>
        /// <returns></returns>
        public bool LoadSettings()
        {
            var configFileSettings = LoadMgrSettingsFromFile();

            return LoadSettings(configFileSettings);
        }

        /// <summary>
        /// Updates manager settings, then loads settings from the database
        /// </summary>
        /// <param name="configFileSettings">Manager settings loaded from file AnalysisManagerProg.exe.config</param>
        /// <returns>True if successful; False on error</returns>
        /// <remarks></remarks>
        public bool LoadSettings(Dictionary<string, string> configFileSettings)
        {
            ErrMsg = string.Empty;

            MgrParams.Clear();

            foreach (var item in configFileSettings)
            {
                MgrParams.Add(item.Key, item.Value);
            }

            // Get directory for main executable
            var appPath = PRISM.FileProcessor.ProcessFilesOrDirectoriesBase.GetAppPath();
            var appFile = new FileInfo(appPath);
            SetParam("ApplicationPath", appFile.DirectoryName);

            // Test the settings retrieved from the config file
            if (!CheckInitialSettings(MgrParams))
            {
                // Error logging handled by CheckInitialSettings
                return false;
            }

            // Assure that MgrActive_Local is defined
            if (!MgrParams.TryGetValue(MGR_PARAM_MGR_ACTIVE_LOCAL, out _))
            {
                // MgrActive_Local parameter not defined defined in the AppName.exe.config file
                HandleParameterNotDefined(MGR_PARAM_MGR_ACTIVE_LOCAL);
            }

            // Get remaining settings from database
            if (!LoadMgrSettingsFromDB())
            {
                // Error logging handled by LoadMgrSettingsFromDB
                return false;
            }

            // Set flag indicating params have been loaded from the manager control database
            ParamsLoadedFromDB = true;

            // No problems found
            return true;
        }

        private Dictionary<string, string> LoadMgrSettingsFromFile()
        {
            // Load initial settings into string dictionary for return
            var mgrSettingsFromFile = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Manager config db connection string
            var mgrCfgDBConnString = Properties.Settings.Default.MgrCnfgDbConnectStr;
            mgrSettingsFromFile.Add(MGR_PARAM_MGR_CFG_DB_CONN_STRING, mgrCfgDBConnString);

            // Manager active flag
            var mgrActiveLocal = Properties.Settings.Default.MgrActive_Local.ToString();
            mgrSettingsFromFile.Add(MGR_PARAM_MGR_ACTIVE_LOCAL, mgrActiveLocal);

            // Manager name
            var mgrName = Properties.Settings.Default.MgrName;
            if (mgrName.Contains("$ComputerName$"))
                mgrSettingsFromFile.Add(MGR_PARAM_MGR_NAME, mgrName.Replace("$ComputerName$", Environment.MachineName));
            else
                mgrSettingsFromFile.Add(MGR_PARAM_MGR_NAME, mgrName);

            // Default settings in use flag
            var usingDefaults = Properties.Settings.Default.UsingDefaults.ToString();
            mgrSettingsFromFile.Add(MGR_PARAM_USING_DEFAULTS, usingDefaults);

            return mgrSettingsFromFile;
        }


        /// <summary>
        /// Tests initial settings retrieved from config file
        /// </summary>
        /// <param name="paramDictionary"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CheckInitialSettings(IReadOnlyDictionary<string, string> paramDictionary)
        {
            // Verify manager settings dictionary exists
            if (paramDictionary == null)
            {
                ErrMsg = "CheckInitialSettings: Manager parameter string dictionary not found";
                OnErrorEvent(ErrMsg);
                return false;
            }

            if (!paramDictionary.TryGetValue(MGR_PARAM_USING_DEFAULTS, out var usingDefaultsText))
            {
                HandleParameterNotDefined(MGR_PARAM_USING_DEFAULTS);
            }
            else
            {
                if (bool.TryParse(usingDefaultsText, out var usingDefaults) && usingDefaults)
                {
                    ErrMsg = string.Format("MgrSettings.CheckInitialSettings; Config file problem, {0} contains UsingDefaults=True",
                                           GetConfigFileName());
                    ReportError(ErrMsg);
                    return false;
                }
            }

            // No problems found
            return true;
        }

        private string GetGroupNameFromSettings(DataTable dtSettings)
        {
            foreach (DataRow currentRow in dtSettings.Rows)
            {
                // Add the column heading and value to the dictionary
                var paramKey = DbCStr(currentRow[dtSettings.Columns["ParameterName"]]);

                if (string.Equals(paramKey, "MgrSettingGroupName", StringComparison.OrdinalIgnoreCase))
                {
                    var groupName = DbCStr(currentRow[dtSettings.Columns["ParameterValue"]]);
                    if (!string.IsNullOrWhiteSpace(groupName))
                    {
                        return groupName;
                    }

                    return string.Empty;
                }
            }

            return string.Empty;
        }

        private void HandleParameterNotDefined(string parameterName)
        {
            ErrMsg = string.Format("Parameter '{0}' is not defined defined in file {1}",
                                   parameterName,
                                   GetConfigFileName());
            ReportError(ErrMsg);
        }

        /// <summary>
        /// Gets manager config settings from manager control DB (Manager_Control)
        /// </summary>
        /// <returns>True if success, otherwise false</returns>
        /// <remarks>Performs retries if necessary.</remarks>
        private bool LoadMgrSettingsFromDB(bool logConnectionErrors = true)
        {

            var managerName = GetParam(MGR_PARAM_MGR_NAME, string.Empty);

            if (string.IsNullOrEmpty(managerName))
            {
                // MgrName parameter not defined defined in the AppName.exe.config file
                HandleParameterNotDefined(MGR_PARAM_MGR_NAME);
                return false;
            }

            var success = LoadMgrSettingsFromDBWork(managerName, out var mgrSettingsFromDB, logConnectionErrors, returnErrorIfNoParameters: true);
            if (!success)
            {
                return false;
            }

            success = StoreParameters(mgrSettingsFromDB, managerName, skipExistingParameters: false);

            while (success)
            {
                var mgrSettingsGroup = GetGroupNameFromSettings(mgrSettingsFromDB);
                if (string.IsNullOrEmpty(mgrSettingsGroup))
                {
                    break;
                }

                // This manager has group-based settings defined; load them now

                success = LoadMgrSettingsFromDBWork(mgrSettingsGroup, out var mgrGroupSettingsFromDB, logConnectionErrors, returnErrorIfNoParameters: false);

                if (success)
                {
                    success = StoreParameters(mgrGroupSettingsFromDB, mgrSettingsGroup, skipExistingParameters: true);
                }
            }

            return success;
        }

        /// <summary>
        /// Load manager settings from the database
        /// </summary>
        /// <param name="managerName">Manager name or manager group name</param>
        /// <param name="mgrSettingsFromDB">Output: manager settings</param>
        /// <param name="logConnectionErrors">When true, log connection errors</param>
        /// <param name="returnErrorIfNoParameters">When true, return an error if no parameters defined</param>
        /// <returns></returns>
        private bool LoadMgrSettingsFromDBWork(
            string managerName,
            out DataTable mgrSettingsFromDB,
            bool logConnectionErrors,
            bool returnErrorIfNoParameters)
        {

            mgrSettingsFromDB = null;

            var dbConnectionString = GetParam(MGR_PARAM_MGR_CFG_DB_CONN_STRING, "");

            if (string.IsNullOrEmpty(dbConnectionString))
            {
                // MgrCnfgDbConnectStr parameter not defined defined in the AppName.exe.config file
                HandleParameterNotDefined(MGR_PARAM_MGR_CFG_DB_CONN_STRING);
                return false;
            }

            var sqlQuery = "SELECT ParameterName, ParameterValue FROM V_MgrParams WHERE ManagerName = '" + managerName + "'";

            // Get a table holding the parameters for this manager
            var retryCount = 3;
            while (retryCount >= 0)
            {
                try
                {
                    using (var cn = new SqlConnection(dbConnectionString))
                    {
                        var cmd = new SqlCommand
                        {
                            CommandType = CommandType.Text,
                            CommandText = sqlQuery,
                            Connection = cn,
                            CommandTimeout = 30
                        };

                        using (var da = new SqlDataAdapter(cmd))
                        {
                            using (var ds = new DataSet())
                            {
                                da.Fill(ds);
                                mgrSettingsFromDB = ds.Tables[0];
                            }
                        }
                    }
                    break;
                }
                catch (Exception ex)
                {
                    retryCount -= 1;
                    var msg = string.Format("LoadMgrSettingsFromDB; Exception getting manager settings from database: {0}; " +
                                            "ConnectionString: {1}, RetryCount = {2}",
                                            ex.Message, dbConnectionString, retryCount);

                    if (logConnectionErrors)
                        ReportError(msg, criticalError: false);

                    // Delay for 5 seconds before trying again
                    if (retryCount >= 0)
                        ProgRunner.SleepMilliseconds(5000);
                }

            } // while

            // If loop exited due to errors, return false
            if (retryCount < 0)
            {
                // Log the message to the DB if the monthly Windows updates are not pending
                var criticalError = !WindowsUpdateStatus.ServerUpdatesArePending();

                ErrMsg = "MgrSettings.LoadMgrSettingsFromDB; Excessive failures attempting to retrieve manager settings from database for manager " + managerName;
                if (logConnectionErrors)
                    ReportError(ErrMsg, criticalError);

                return false;
            }

            // Validate that the data table object is initialized
            if (mgrSettingsFromDB == null)
            {
                // Data table not initialized
                ErrMsg = "LoadMgrSettingsFromDB; dtSettings data table is null; using " + dbConnectionString;
                if (logConnectionErrors)
                    ReportError(ErrMsg);

                return false;
            }

            // Verify at least one row returned
            if (mgrSettingsFromDB.Rows.Count < 1 && returnErrorIfNoParameters)
            {
                // Wrong number of rows returned
                ErrMsg = string.Format("MgrSettings.LoadMgrSettingsFromDB; Manager {0} not defined in the manager control database; using {1}",
                                       managerName, dbConnectionString);
                ReportError(ErrMsg);
                mgrSettingsFromDB.Dispose();
                return false;
            }

            return true;
        }

        /// <summary>
        /// Update mParamDictionary with settings in dtSettings, optionally skipping existing parameters
        /// </summary>
        /// <param name="dtSettings"></param>
        /// <param name="managerOrGroupName"></param>
        /// <param name="skipExistingParameters"></param>
        /// <returns></returns>
        private bool StoreParameters(DataTable dtSettings, string managerOrGroupName, bool skipExistingParameters)
        {
            bool success;

            try
            {
                foreach (DataRow currentRow in dtSettings.Rows)
                {
                    // Add the column heading and value to the dictionary
                    var paramKey = DbCStr(currentRow[dtSettings.Columns["ParameterName"]]);
                    var paramVal = DbCStr(currentRow[dtSettings.Columns["ParameterValue"]]);

                    if (paramKey.ToLower() == "perspective" && Environment.MachineName.ToLower().StartsWith("monroe"))
                    {
                        if (paramVal.ToLower() == "server")
                        {
                            paramVal = "client";
                            Console.WriteLine(
                                @"StoreParameters: Overriding manager perspective to be 'client' because impersonating a server-based manager from an office computer");
                        }
                    }

                    if (MgrParams.ContainsKey(paramKey))
                    {
                        if (!skipExistingParameters)
                        {
                            MgrParams[paramKey] = paramVal;
                        }
                    }
                    else
                    {
                        MgrParams.Add(paramKey, paramVal);
                    }
                }
                success = true;
            }
            catch (Exception ex)
            {
                ErrMsg = string.Format("MgrSettings.StoreParameters; Exception storing settings for manager '{0}': {1}",
                                       managerOrGroupName, ex.Message);
                ReportError(ErrMsg);
                success = false;
            }
            finally
            {
                dtSettings?.Dispose();
            }

            return success;
        }

        /// <summary>
        /// Gets a manager parameter
        /// </summary>
        /// <param name="itemKey"></param>
        /// <returns>Parameter value if found, otherwise empty string</returns>
        public string GetParam(string itemKey)
        {
            return GetParam(itemKey, string.Empty);
        }

        /// <summary>
        /// Gets a manager parameter
        /// </summary>
        /// <param name="itemKey">Parameter name</param>
        /// <param name="valueIfMissing">Value to return if the parameter does not exist</param>
        /// <returns>Parameter value if found, otherwise empty string</returns>
        public string GetParam(string itemKey, string valueIfMissing)
        {
            if (MgrParams.TryGetValue(itemKey, out var itemValue))
            {
                return itemValue ?? string.Empty;
            }

            return valueIfMissing ?? string.Empty;
        }

        /// <summary>
        /// Gets a manager parameter
        /// </summary>
        /// <param name="itemKey">Parameter name</param>
        /// <param name="valueIfMissing">Value to return if the parameter does not exist</param>
        /// <returns>Parameter value if found, otherwise empty string</returns>
        public bool GetParam(string itemKey, bool valueIfMissing)
        {
            if (MgrParams.TryGetValue(itemKey, out var valueText))
            {
                var value = UtilityMethods.CBoolSafe(valueText, valueIfMissing);
                return value;
            }

            return valueIfMissing;
        }

        /// <summary>
        /// Gets a manager parameter
        /// </summary>
        /// <param name="itemKey">Parameter name</param>
        /// <param name="valueIfMissing">Value to return if the parameter does not exist</param>
        /// <returns>Parameter value if found, otherwise empty string</returns>
        public int GetParam(string itemKey, int valueIfMissing)
        {
            if (MgrParams.TryGetValue(itemKey, out var valueText))
            {
                var value = UtilityMethods.CIntSafe(valueText, valueIfMissing);
                return value;
            }

            return valueIfMissing;
        }

        /// <summary>
        /// Set a manager parameter
        /// </summary>
        /// <param name="itemKey"></param>
        /// <param name="itemValue"></param>
        // ReSharper disable once UnusedMember.Global
        public void SetParam(string itemKey, string itemValue)
        {
            if (MgrParams.ContainsKey(itemKey))
            {
                MgrParams[itemKey] = itemValue;
            }
            else
            {
                MgrParams.Add(itemKey, itemValue);
            }
        }

        private string DbCStr(object inpObj)
        {
            if (inpObj == null)
            {
                return "";
            }
            return inpObj.ToString();
        }

        #endregion

        #region "Event Handlers"

        /// <summary>
        /// Report an important error
        /// </summary>
        /// <param name="message"></param>
        private void OnCriticalErrorEvent(string message)
        {
            if (CriticalErrorEvent == null && WriteToConsoleIfNoListener)
                ConsoleMsgUtils.ShowError(message, false, false, EmptyLinesBeforeErrorMessages);

            CriticalErrorEvent?.Invoke(message, null);
        }

        /// <summary>
        /// Raises a CriticalErrorEvent if criticalError is true and ParamsLoadedFromDB is false
        /// Otherwise, raises a normal error event
        /// </summary>
        /// <param name="errorMessage"></param>
        /// <param name="criticalError"></param>
        private void ReportError(string errorMessage, bool criticalError = true)
        {
            if (!ParamsLoadedFromDB && criticalError)
            {
                OnCriticalErrorEvent(errorMessage);
            }
            else
            {
                OnErrorEvent(errorMessage);
            }
        }

        #endregion
    }
}
