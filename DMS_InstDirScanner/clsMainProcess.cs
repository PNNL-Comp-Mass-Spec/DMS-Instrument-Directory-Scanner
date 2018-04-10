'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
' This is a complete rewrite of the Instrument Directory Scanner program written by Nathan Trimble
'
'*********************************************************************************************************

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports DMS_InstDirScanner.My
Imports PRISM
Imports PRISM.Logging

''' <summary>
''' Master processing class
''' </summary>
Public Class clsMainProcess

#Region "Constants"

    Private Const DEFAULT_BASE_LOGFILE_NAME = "Logs\InstDirScanner"

    Private Const MGR_PARAM_DEFAULT_DMS_CONN_STRING = "MgrCnfgDbConnectStr"

#End Region

#Region "Module variables"
    Shared m_MainProcess As clsMainProcess

    Private ReadOnly m_MgrExeName As String

    Private ReadOnly m_MgrDirectoryPath As String

    Private m_MgrSettings As clsMgrSettings

    Shared m_StatusFile As clsStatusFile

#End Region

#Region "Methods"

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        Dim exeInfo = New FileInfo(FileProcessor.ProcessFilesOrFoldersBase.GetAppPath())
        m_MgrExeName = exeInfo.Name
        m_MgrDirectoryPath = exeInfo.DirectoryName
    End Sub

    ''' <summary>
    ''' Starts program execution
    ''' </summary>
    ''' <remarks></remarks>
    Shared Sub Main()

        Try
            If IsNothing(m_MainProcess) Then
                m_MainProcess = New clsMainProcess()
                If Not m_MainProcess.InitMgr() Then
                    Thread.Sleep(1500)
                    Exit Sub
                End If
            End If
            m_MainProcess.DoDirectoryScan()
        Catch ex As Exception
            ' Report any exceptions not handled at a lower level to the system application log
            Const errMsg = "Critical exception starting application"
            LogTools.LogError(errMsg, ex, True)
            Thread.Sleep(1500)
            Exit Sub
        Finally
            If Not m_StatusFile Is Nothing Then
                m_StatusFile.DisposeMessageQueue()
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Initializes the manager settings and classes
    ''' </summary>
    ''' <returns>True for success; False if error occurs</returns>
    ''' <remarks></remarks>
    Private Function InitMgr() As Boolean

        Dim hostName = Net.Dns.GetHostName()

        ' Define the default logging info
        ' This will get updated below
        LogTools.CreateFileLogger(DEFAULT_BASE_LOGFILE_NAME, BaseLogger.LogLevels.DEBUG)

        ' Create a database logger connected to DMS5
        ' Once the initial parameters have been successfully read,
        ' we update the dbLogger to use the connection string read from the Manager Control DB
        Dim defaultDmsConnectionString As String

        ' Open DMS_InstDirScanner.exe.config to look for setting DefaultDMSConnString, so we know which server to log to by default
        Dim dmsConnectionStringFromConfig = GetXmlConfigDefaultConnectionString()

        If String.IsNullOrWhiteSpace(dmsConnectionStringFromConfig) Then
            ' Use the hard-coded default that points to Proteinseqs
            defaultDmsConnectionString = MySettings.Default.MgrCnfgDbConnectStr
        Else
            ' Use the connection string from DMS_InstDirScanner.exe.config
            defaultDmsConnectionString = dmsConnectionStringFromConfig
        End If

        ConsoleMsgUtils.ShowDebug("Instantiate a DbLogger using " + defaultDmsConnectionString)

        LogTools.CreateDbLogger(defaultDmsConnectionString, "CaptureTaskMan: " + hostName)

        ' Get the manager settings
        Try
            m_MgrSettings = New clsMgrSettings()
        Catch ex As Exception
            ' Failures are logged by clsMgrSettings to application event logs
            Return False
        End Try

        ' Setup the logger
        Dim logFileNameBase As String = m_MgrSettings.GetParam("logfilename", "InstDirScanner")

        Dim debugLevel As Integer = m_MgrSettings.GetParam("debuglevel", 1)

        Dim logLevel = CType(debugLevel, BaseLogger.LogLevels)
        LogTools.CreateFileLogger(logFileNameBase, logLevel)

        Dim logCnStr As String = m_MgrSettings.GetParam("connectionstring")
        Dim moduleName As String = m_MgrSettings.GetParam("modulename")
        LogTools.CreateDbLogger(logCnStr, moduleName)

        ' Make the initial log entry
        Dim myMsg As String = "=== Started Instrument Directory Scanner V" & GetAppVersion() & " ===== "
        LogTools.LogMessage(myMsg)

        ' Setup the status file class
        Dim statusFileNameLoc As String = Path.Combine(GetAppFolderPath(), "Status.xml")
        m_StatusFile = New clsStatusFile(statusFileNameLoc, debugLevel)
        AttachEvents(m_StatusFile)

        With m_StatusFile
            .MessageQueueURI = m_MgrSettings.GetParam("MessageQueueURI")
            .MessageQueueTopic = m_MgrSettings.GetParam("MessageQueueTopicMgrStatus")
            .LogToMsgQueue = CBool(m_MgrSettings.GetParam("LogStatusToMessageQueue"))
            .MgrName = m_MgrSettings.GetParam("MgrName")
            .MgrStatus = IStatusFile.EnumMgrStatus.Running
            .WriteStatusFile()
        End With

        ' Everything worked
        Return True

    End Function

    ''' <summary>
    ''' Do a directory scan
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DoDirectoryScan()

        Try

            ' Check to see if manager is active
            If Not CBool(m_MgrSettings.GetParam("mgractive")) Then
                Dim message = "Program disabled in manager control DB"
                ConsoleMsgUtils.ShowWarning(message)
                LogTools.LogMessage(message)
                LogTools.LogMessage("===== Closing Inst Dir Scanner =====")
                m_StatusFile.UpdateDisabled(False)
                Exit Sub
            ElseIf Not CBool(m_MgrSettings.GetParam("mgractive_local")) Then
                Dim message = "Program disabled locally"
                ConsoleMsgUtils.ShowWarning(message)
                LogTools.LogMessage(message)
                LogTools.LogMessage("===== Closing Inst Dir Scanner =====")
                m_StatusFile.UpdateDisabled(True)
                Exit Sub
            End If

            Dim workDir = m_MgrSettings.GetParam("workdir", String.Empty)
            If String.IsNullOrWhiteSpace(workDir) Then
                LogFatalError("Manager parameter 'workdir' is not defined")
                Exit Sub
            End If

            ' Verify output directory can be found
            If Not Directory.Exists(workDir) Then
                LogFatalError("Output directory not found: " & workDir)
                Exit Sub
            End If

            ' Get list of instruments from DMS
            Dim instList As List(Of clsInstData) = GetInstrumentList()
            If instList Is Nothing Then
                LogFatalError("No instrument list")
                Exit Sub
            End If

            ' Scan the directories
            Dim scanner = New clsDirectoryTools()
            AttachEvents(scanner)

            scanner.PerformDirectoryScans(instList, workDir, m_MgrSettings, m_StatusFile)

            ' All finished, so clean up and exit
            LogMessage("Scanning complete")
            LogTools.LogMessage("===== Closing Inst Dir Scanner =====")
            m_StatusFile.UpdateStopped(False)

        Catch ex As Exception
            LogError("Error in DoDirectoryScan", ex)
        End Try

    End Sub

    Public Shared Function GetAppFolderPath() As String
        ' Could use Application.StartupPath, but .GetExecutingAssembly is better
        Return Path.GetDirectoryName(GetAppPath())
    End Function

    ''' <summary>
    ''' Returns the full path to the executing .Exe or .Dll
    ''' </summary>
    ''' <returns>File path</returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppPath() As String
        Return Assembly.GetExecutingAssembly().Location
    End Function

    ''' <summary>
    ''' Returns the .NET assembly version followed by the program date
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppVersion() As String
        Return Assembly.GetExecutingAssembly().GetName().Version.ToString()
    End Function

    Private Function GetInstrumentList() As List(Of clsInstData)

        LogMessage("Getting instrument list")

        Dim columns = New List(Of String) From {
            "vol",
            "path",
            "method",
            "Instrument"
        }

        Dim sqlQuery = "SELECT " + String.Join(",", columns) + " FROM V_Instrument_Source_Paths"

        Dim connectionString = m_MgrSettings.GetParam("connectionstring")
        If String.IsNullOrWhiteSpace(connectionString) Then
            LogError("Connection string is empty; cannot retrieve manager parameters")
            Return Nothing
        End If

        Dim dbTools = New clsDBTools(connectionString)
        AttachEvents(dbTools)

        ' Get a table containing the active instruments
        Dim lstResults As List(Of List(Of String)) = Nothing
        Dim success = dbTools.GetQueryResults(sqlQuery, lstResults, "GetInstrumentList")

        ' Verify valid data found
        If Not success OrElse lstResults Is Nothing Then
            LogError("Unable to retrieve instrument list")
            Return Nothing
        End If

        If lstResults.Count < 1 Then
            LogError("No instruments found")
            Return Nothing
        End If

        Dim colMapping = dbTools.GetColumnMapping(columns)

        ' Create a list of all instrument data
        Dim instrumentList As New List(Of clsInstData)
        Try
            For Each result In lstResults
                Dim instrumentInfo As New clsInstData With {
                    .CaptureMethod = dbTools.GetColumnValue(result, colMapping, "method"),
                    .InstName = dbTools.GetColumnValue(result, colMapping, "Instrument"),
                    .StoragePath = dbTools.GetColumnValue(result, colMapping, "path"),
                    .StorageVolume = dbTools.GetColumnValue(result, colMapping, "vol")
                }
                instrumentList.Add(instrumentInfo)
            Next
            LogTools.LogDebug("Retrieved instrument list")
            Return instrumentList

        Catch ex As Exception
            LogError("Exception filling instrument list", ex)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Extract the value DefaultDMSConnString from DMS_InstDirScanner.exe.config
    ''' </summary>
    ''' <returns></returns>
    Private Function GetXmlConfigDefaultConnectionString() As String
        Return GetXmlConfigFileSetting(MGR_PARAM_DEFAULT_DMS_CONN_STRING)
    End Function

    ''' <summary>
    ''' Extract the value for the given setting from DMS_InstDirScanner.exe.config
    ''' </summary>
    ''' <returns>Setting value if found, otherwise an empty string</returns>
    ''' <remarks>Uses a simple text reader in case the file has malformed XML</remarks>
    Private Function GetXmlConfigFileSetting(settingName As String) As String

        If String.IsNullOrWhiteSpace(settingName) Then
            Throw New ArgumentException("Setting name cannot be blank", NameOf(settingName))
        End If

        Try
            Dim configFilePath = Path.Combine(m_MgrDirectoryPath, m_MgrExeName + ".config")
            Dim configfile = New FileInfo(configFilePath)
            If Not configfile.Exists Then
                LogError("File not found: " + configFilePath)
                Return String.Empty
            End If

            Dim configXml = New StringBuilder()

            ' Open DMS_InstDirScanner.exe.config using a simple text reader in case the file has malformed XML
            ConsoleMsgUtils.ShowDebug(String.Format("Extracting setting {0} from {1}", settingName, configfile.FullName))

            Using reader = New StreamReader(New FileStream(configfile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine
                    If String.IsNullOrWhiteSpace(dataLine) Then
                        Continue While
                    End If

                    configXml.Append(dataLine)
                End While
            End Using

            Dim matcher = New Regex((settingName + ".+?<value>(?<ConnString>.+?)</value>"), RegexOptions.IgnoreCase)
            Dim match = matcher.Match(configXml.ToString)
            If match.Success Then
                Return match.Groups("ConnString").Value
            End If

            LogError(settingName + " setting not found in " + configFilePath)
            Return String.Empty

        Catch ex As Exception
            LogError("Exception reading setting " + settingName + " in DMS_InstDirScanner.exe.config", ex)
            Return String.Empty
        End Try

    End Function
#End Region

#Region "Event handlers"

    Private Sub AttachEvents(objClass As clsEventNotifier)
        AddHandler objClass.ErrorEvent, AddressOf ErrorHandler
        AddHandler objClass.StatusEvent, AddressOf MessageHandler
        AddHandler objClass.WarningEvent, AddressOf WarningHandler
    End Sub

    Private Sub LogError(message As String, Optional ex As Exception = Nothing)
        LogTools.LogError(message, ex)
    End Sub

    Private Sub LogMessage(message As String)
        LogTools.LogMessage(message)
    End Sub

    Private Sub LogWarning(message As String)
        LogTools.LogWarning(message)
    End Sub

    Private Sub LogFatalError(errorMessage As String)
        LogError(errorMessage)
        LogTools.LogMessage("===== Closing Inst Dir Scanner =====")
        m_StatusFile.UpdateStopped(True)
    End Sub

    Private Sub ErrorHandler(message As String, ex As Exception)
        LogError(message)
    End Sub

    Private Sub MessageHandler(message As String)
        LogTools.LogMessage(message)
    End Sub

    Private Sub WarningHandler(message As String)
        LogWarning(message)
    End Sub
#End Region

End Class
