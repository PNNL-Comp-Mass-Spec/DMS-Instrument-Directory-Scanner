﻿'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
' This is a complete rewrite of the Instrument Directory Scanner program written by Nathan Trimble
'
' Last modified 07/27/2009
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
	Private m_StatusFile As IStatusFile
#End Region

#Region "Methods"
	''' <summary>
	''' Starts program execution
	''' </summary>
	''' <remarks></remarks>
	Shared Sub Main()

		Dim ErrMsg As String

		Try
			If IsNothing(m_MainProcess) Then
				m_MainProcess = New clsMainProcess
				If Not m_MainProcess.InitMgr Then Exit Sub
			End If
			m_MainProcess.DoDirectoryScan()
		Catch Err As Exception
			'Report any exceptions not handled at a lower level to the system application log
			ErrMsg = "Critical exception starting application"
			WriteLog(LoggerTypes.LogSystem, LogLevels.FATAL, ErrMsg, Err)
			Exit Sub
		End Try

	End Sub

	''' <summary>
	''' Constructor
	''' </summary>
	''' <remarks>Doesn't do anything at present</remarks>
	Public Sub New()

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
		Dim LogFileName As String = m_MgrSettings.GetParam("logfilename")
		Dim DebugLevel As Integer = CInt(m_MgrSettings.GetParam("debuglevel"))
		clsLogTools.CreateFileLogger(LogFileName, DebugLevel)
		Dim LogCnStr As String = m_MgrSettings.GetParam("connectionstring")
		Dim ModuleName As String = m_MgrSettings.GetParam("modulename")
		clsLogTools.CreateDbLogger(LogCnStr, ModuleName)

		'Make the initial log entry
		Dim MyMsg As String = "=== Started Instrument Directory Scanner V" & Application.ProductVersion & " ===== "
		WriteLog(LoggerTypes.LogFile, LogLevels.INFO, MyMsg)

		'Setup the status file class
		Dim FInfo As FileInfo = New FileInfo(Application.ExecutablePath)
		Dim StatusFileNameLoc As String = Path.Combine(FInfo.DirectoryName, "Status.xml")
		m_StatusFile = New clsStatusFile(StatusFileNameLoc)
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

		'Verify output directory can be found
		If Not Directory.Exists(m_MgrSettings.GetParam("workdir")) Then
			WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Output directory not found")
			WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
			m_StatusFile.UpdateStopped(True)
			Exit Sub
		End If

		'Get list of instruments from DMS
		Dim InstList As List(Of clsInstData) = clsDbTools.GetInstrumentList(m_MgrSettings)
		If InstList Is Nothing Then
			WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "No instrument list")
			WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
			m_StatusFile.UpdateStopped(True)
			Exit Sub
		End If

		'Scan the directories
		clsDirectoryTools.PerformDirectoryScans(InstList, m_MgrSettings.GetParam("workdir"), m_MgrSettings, m_StatusFile)

		'All finished, so clean up and exit
		WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Scanning complete")
		WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "===== Closing Inst Dir Scanner =====")
		m_StatusFile.UpdateStopped(False)

	End Sub
#End Region

End Class