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
Imports DMS_InstDirScanner.MgrSettings
Imports DMS_InstDirScanner.clsLogTools

Public Class clsMainProcess

	'*********************************************************************************************************
	' Master processing class
	'*********************************************************************************************************

#Region "Module variables"
	Shared m_MainProcess As clsMainProcess
	Private m_MgrSettings As clsMgrSettings
	Shared m_StatusFile As IStatusFile
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
                If Not m_MainProcess.InitMgr Then Exit Sub
            End If
            m_MainProcess.DoDirectoryScan()
        Catch ex As Exception
            'Report any exceptions not handled at a lower level to the system application log
            Const errMsg As String = "Critical exception starting application"
            WriteLog(LoggerTypes.LogDb, LogLevels.FATAL, errMsg, ex)
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

		'Get the manager settings
		Try
			m_MgrSettings = New clsMgrSettings()
		Catch ex As Exception
			'Failures are logged by clsMgrSettings to application event logs
			Return False
		End Try

		'Setup the logger
        Dim logFileName As String = m_MgrSettings.GetParam("logfilename")
        Dim debugLevel As Integer = m_MgrSettings.GetParam("debuglevel", 1)
        clsLogTools.CreateFileLogger(logFileName, debugLevel)

        Dim logCnStr As String = m_MgrSettings.GetParam("connectionstring")
        Dim moduleName As String = m_MgrSettings.GetParam("modulename")
        clsLogTools.CreateDbLogger(logCnStr, moduleName)

		'Make the initial log entry
        Dim myMsg As String = "=== Started Instrument Directory Scanner V" & GetAppVersion() & " ===== "
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, myMsg)

		'Setup the status file class
        Dim statusFileNameLoc As String = Path.Combine(GetAppFolderPath(), "Status.xml")
        m_StatusFile = New clsStatusFile(statusFileNameLoc, debugLevel)

		With m_StatusFile
			.MessageQueueURI = m_MgrSettings.GetParam("MessageQueueURI")
			.MessageQueueTopic = m_MgrSettings.GetParam("MessageQueueTopicMgrStatus")
			.LogToMsgQueue = CBool(m_MgrSettings.GetParam("LogStatusToMessageQueue"))
			.MgrName = m_MgrSettings.GetParam("MgrName")
			.MgrStatus = IStatusFile.EnumMgrStatus.Running
			.WriteStatusFile()
		End With

		'Everything worked
		Return True

	End Function

	''' <summary>
	''' Do a directory scan
	''' </summary>
	''' <remarks></remarks>
	Private Sub DoDirectoryScan()

        Try

            'Check to see if manager is active
            If Not CBool(m_MgrSettings.GetParam("mgractive")) Then
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Program disabled in manager control DB")
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
                m_StatusFile.UpdateDisabled(False)
                Exit Sub
            ElseIf Not CBool(m_MgrSettings.GetParam("mgractive_local")) Then
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Program disabled locally")
                WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
                m_StatusFile.UpdateDisabled(True)
                Exit Sub
            End If

            Dim workDir = m_MgrSettings.GetParam("workdir", String.Empty)
            If String.IsNullOrWhiteSpace(workDir) Then
                LogFatalError("Manager parameter 'workdir' is not defined")
                Exit Sub
            End If

            'Verify output directory can be found
            If Not Directory.Exists(workDir) Then
                LogFatalError("Output directory not found: " & workDir)
                Exit Sub
            End If

            'Get list of instruments from DMS
            Dim instList As List(Of clsInstData) = clsDbTools.GetInstrumentList(m_MgrSettings)
            If instList Is Nothing Then
                LogFatalError("No instrument list")
                Exit Sub
            End If

            'Scan the directories
            Dim scanner = New clsDirectoryTools()
            scanner.PerformDirectoryScans(instList, workDir, m_MgrSettings, m_StatusFile)

            'All finished, so clean up and exit
            WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Scanning complete")
            WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
            m_StatusFile.UpdateStopped(False)

        Catch ex As Exception
            WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Error in DoDirectoryScan: " & ex.Message)
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
        Return Reflection.Assembly.GetExecutingAssembly().Location
    End Function

    ''' <summary>
    ''' Returns the .NET assembly version followed by the program date
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetAppVersion() As String
        Return Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
    End Function

    Private Sub LogFatalError(ByVal errorMessage As String)
        WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, errorMessage)
        WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
        m_StatusFile.UpdateStopped(True)
    End Sub

#End Region

End Class
