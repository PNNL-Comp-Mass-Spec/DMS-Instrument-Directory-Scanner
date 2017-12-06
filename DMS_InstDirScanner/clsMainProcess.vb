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
Imports System.Threading
Imports DMS_InstDirScanner.clsLogTools
Imports PRISM

''' <summary>
''' Master processing class
''' </summary>
Public Class clsMainProcess

#Region "Module variables"
    Shared m_MainProcess As clsMainProcess
    Private m_MgrSettings As clsMgrSettings
    Shared m_StatusFile As clsStatusFile
#End Region

#Region "Methods"
    ''' <summary>
    ''' Starts program execution
    ''' </summary>
    ''' <remarks></remarks>
    Shared Sub Main()

        Try
            If IsNothing(m_MainProcess) Then
                m_MainProcess = New clsMainProcess
                If Not m_MainProcess.InitMgr() Then
                    Thread.Sleep(1500)
                    Exit Sub
                End If
            End If
            m_MainProcess.DoDirectoryScan()
        Catch ex As Exception
            ' Report any exceptions not handled at a lower level to the system application log
            Const errMsg = "Critical exception starting application"
            ConsoleMsgUtils.ShowError(errMsg)
            WriteLog(LoggerTypes.LogDb, LogLevels.FATAL, errMsg, ex)
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

        ' Get the manager settings
        Try
            m_MgrSettings = New clsMgrSettings()
        Catch ex As Exception
            ' Failures are logged by clsMgrSettings to application event logs
            Return False
        End Try

        ' Setup the logger
        Dim logFileNameBase As String = m_MgrSettings.GetParam("logfilename")
        If String.IsNullOrWhiteSpace(logFileNameBase) Then
            logFileNameBase = "InstDirScanner"
        End If

        Dim debugLevel As Integer = m_MgrSettings.GetParam("debuglevel", 1)
        CreateFileLogger(logFileNameBase, debugLevel)

        Dim logCnStr As String = m_MgrSettings.GetParam("connectionstring")
        Dim moduleName As String = m_MgrSettings.GetParam("modulename")
        CreateDbLogger(logCnStr, moduleName)

        ' Make the initial log entry
        Dim myMsg As String = "=== Started Instrument Directory Scanner V" & GetAppVersion() & " ===== "
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, myMsg)

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
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, message)
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
                m_StatusFile.UpdateDisabled(False)
                Exit Sub
            ElseIf Not CBool(m_MgrSettings.GetParam("mgractive_local")) Then
                Dim message = "Program disabled locally"
                ConsoleMsgUtils.ShowWarning(message)
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, message)
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
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
            WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
            m_StatusFile.UpdateStopped(False)

        Catch ex As Exception
            LogError("Error in DoDirectoryScan: " & ex.Message)
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
            WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, "Retrieved instrument list")
            Return instrumentList

        Catch ex As Exception
            LogError("Exception filling instrument list: " & ex.Message)
            Return Nothing
        End Try

    End Function

#End Region

#Region "Event handlers"

    Private Sub AttachEvents(objClass As clsEventNotifier)
        AddHandler objClass.ErrorEvent, AddressOf ErrorHandler
        AddHandler objClass.StatusEvent, AddressOf MessageHandler
        AddHandler objClass.WarningEvent, AddressOf WarningHandler
    End Sub

    Private Sub LogError(message As String)
        ConsoleMsgUtils.ShowError(message)
        WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, message)
    End Sub

    Private Sub LogMessage(message As String)
        Console.WriteLine(message)
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, message)
    End Sub

    Private Sub LogWarning(message As String)
        ConsoleMsgUtils.ShowWarning(message)
        WriteLog(LoggerTypes.LogFile, LogLevels.WARN, message)
    End Sub

    Private Sub LogFatalError(errorMessage As String)
        LogError(errorMessage)
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
        m_StatusFile.UpdateStopped(True)
    End Sub

    Private Sub ErrorHandler(message As String, ex As Exception)
        LogError(message)
    End Sub

    Private Sub MessageHandler(message As String)
        Console.WriteLine(message)
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, message)
    End Sub

    Private Sub WarningHandler(message As String)
        LogWarning(message)
    End Sub
#End Region

End Class
