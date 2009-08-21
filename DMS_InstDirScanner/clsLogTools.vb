'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
' Last modified 07/27/2009
'						- 08/21/2009 (DAC) - Added additional status parameters
'*********************************************************************************************************
Imports log4net
Imports System.Data

'This assembly attribute tells Log4Net where to find the config file
<Assembly: log4net.Config.XmlConfigurator(ConfigFile:="Logging.config", Watch:=True)> 

Public Class clsLogTools

	'*********************************************************************************************************
	' Class for handling logging via Log4Net
	'*********************************************************************************************************

#Region "Enums"
	Public Enum LogLevels
		DEBUG = 5
		INFO = 4
		WARN = 3
		[ERROR] = 2
		FATAL = 1
	End Enum

	Public Enum LoggerTypes
		LogFile
		LogDb
		LogSystem
	End Enum
#End Region

#Region "Module variables"
	Private Shared ReadOnly m_FileLogger As ILog = LogManager.GetLogger("FileLogger")
	Private Shared ReadOnly m_DbLogger As ILog = LogManager.GetLogger("DbLogger")
	Private Shared ReadOnly m_SysLogger As ILog = LogManager.GetLogger("SysLogger")
	Private Shared m_FileDate As String
	Private Shared m_BaseFileName As String
	Private Shared m_FileAppender As log4net.Appender.FileAppender
#End Region

#Region "Properties"
	''' <summary>
	''' Tells calling program file debug status
	''' </summary>
	''' <returns>TRUE if debug level enabled for file logger; FALSE otherwise</returns>
	''' <remarks></remarks>
	Public Shared ReadOnly Property FileLogDebugEnabled() As Boolean
		Get
			Return m_FileLogger.IsDebugEnabled
		End Get
	End Property
#End Region

#Region "Methods"
	''' <summary>
	''' Writes a message to the logging system
	''' </summary>
	''' <param name="LoggerType">Type of logger to use</param>
	''' <param name="LogLevel">Level of log reporting</param>
	''' <param name="InpMsg">Message to be logged</param>
	''' <remarks></remarks>
	Public Shared Sub WriteLog(ByVal LoggerType As LoggerTypes, ByVal LogLevel As LogLevels, ByVal InpMsg As String)

		Dim MyLogger As ILog

		'Establish which logger will be used
		Select Case LoggerType
			Case LoggerTypes.LogDb
				MyLogger = m_DbLogger
			Case LoggerTypes.LogFile
				MyLogger = m_FileLogger
				Dim TestFileDate As String = Now.ToString("MM-dd-yyyy")
				If TestFileDate <> m_FileDate Then
					m_FileDate = TestFileDate
					ChangeLogFileName()
				End If
			Case LoggerTypes.LogSystem
				MyLogger = m_SysLogger
			Case Else
				Throw New Exception("Invalid logger type specified")
		End Select

		'Update the status file data
		clsStatusData.MostRecentLogMessage = Now.ToString("MM/dd/yyyy HH:mm:ss") & "; " & InpMsg _
			  & "; " & LogLevel.ToString()
		'Send the log message
		Select Case LogLevel
			Case LogLevels.DEBUG
				If MyLogger.IsDebugEnabled Then MyLogger.Debug(InpMsg)
			Case LogLevels.ERROR
				clsStatusData.AddErrorMessage(Now.ToString("MM/dd/yyyy HH:mm:ss") & "; " & InpMsg _
					 & "; " & LogLevel.ToString())
				If MyLogger.IsErrorEnabled Then MyLogger.Error(InpMsg)
			Case LogLevels.FATAL
				If MyLogger.IsFatalEnabled Then MyLogger.Fatal(InpMsg)
			Case LogLevels.INFO
				If MyLogger.IsInfoEnabled Then MyLogger.Info(InpMsg)
			Case LogLevels.WARN
				If MyLogger.IsWarnEnabled Then MyLogger.Warn(InpMsg)
			Case Else
				Throw New Exception("Invalid log level specified")
		End Select
	End Sub

	''' <summary>
	''' Overload to write a message and exception to the logging system
	''' </summary>
	''' <param name="LoggerType">Type of logger to use</param>
	''' <param name="LogLevel">Level of log reporting</param>
	''' <param name="InpMsg">Message to be logged</param>
	''' <param name="Ex">Exception to be logged</param>
	''' <remarks></remarks>
	Public Shared Sub WriteLog(ByVal LoggerType As LoggerTypes, ByVal LogLevel As LogLevels, ByVal InpMsg As String, _
	 ByVal Ex As Exception)

		Dim MyLogger As ILog

		'Establish which logger will be used
		Select Case LoggerType
			Case LoggerTypes.LogDb
				MyLogger = m_DbLogger
			Case LoggerTypes.LogFile
				MyLogger = m_FileLogger
				Dim TestFileDate As String = Now.ToString("MM-dd-yyyy")
				If TestFileDate <> m_FileDate Then
					m_FileDate = TestFileDate
					ChangeLogFileName()
				End If
			Case LoggerTypes.LogSystem
				MyLogger = m_SysLogger
			Case Else
				Throw New Exception("Invalid logger type specified")
		End Select

		'Update the status file data
		clsStatusData.MostRecentLogMessage = Now.ToString("MM/dd/yyyy HH:mm:ss") & "; " & InpMsg _
			 & "; " & LogLevel.ToString()
		'Send the log message
		Select Case LogLevel
			Case LogLevels.DEBUG
				If MyLogger.IsDebugEnabled Then MyLogger.Debug(InpMsg, Ex)
			Case LogLevels.ERROR
				clsStatusData.AddErrorMessage(Now.ToString("MM/dd/yyyy HH:mm:ss") & "; " & InpMsg _
					 & "; " & LogLevel.ToString)
				If MyLogger.IsErrorEnabled Then MyLogger.Error(InpMsg, Ex)
			Case LogLevels.FATAL
				If MyLogger.IsFatalEnabled Then MyLogger.Fatal(InpMsg, Ex)
			Case LogLevels.INFO
				If MyLogger.IsInfoEnabled Then MyLogger.Info(InpMsg, Ex)
			Case LogLevels.WARN
				If MyLogger.IsWarnEnabled Then MyLogger.Warn(InpMsg, Ex)
			Case Else
				Throw New Exception("Invalid log level specified")
		End Select
	End Sub

	''' <summary>
	''' Changes the base log file name
	''' </summary>
	''' <remarks></remarks>
	Public Shared Sub ChangeLogFileName()

		'Get a list of appenders
		Dim AppendList As List(Of Appender.IAppender) = FindAppenders("FileAppender")
		If AppendList Is Nothing Then
			WriteLog(LoggerTypes.LogSystem, LogLevels.WARN, "Unable to change file name. No appender found")
			Return
		End If

		For Each SelectedAppender As Appender.IAppender In AppendList
			'Convert the IAppender object to a RollingFileAppender
			Dim AppenderToChange As Appender.FileAppender = TryCast(SelectedAppender, Appender.FileAppender)
			If AppenderToChange Is Nothing Then
				WriteLog(LoggerTypes.LogSystem, LogLevels.ERROR, "Unable to convert appender")
				Return
			End If
			'Change the file name and activate change
			AppenderToChange.File = m_BaseFileName & "_" & m_FileDate & ".txt"
			AppenderToChange.ActivateOptions()
		Next
	End Sub

	''' <summary>
	''' Gets the specified appender
	''' </summary>
	''' <param name="AppendName">Name of appender to find</param>
	''' <returns>List(IAppender) objects if found; NOTHING otherwise</returns>
	''' <remarks></remarks>
	Private Shared Function FindAppenders(ByVal AppendName As String) As List(Of Appender.IAppender)

		'Get a list of the current loggers
		Dim LoggerList() As ILog = LogManager.GetCurrentLoggers()
		If LoggerList.GetLength(0) < 1 Then Return Nothing

		'Create a List of appenders matching the criteria for each logger
		Dim RetList As New List(Of Appender.IAppender)
		For Each TestLogger As ILog In LoggerList
			For Each TestAppender As Appender.IAppender In TestLogger.Logger.Repository.GetAppenders()
				If TestAppender.Name = AppendName Then RetList.Add(TestAppender)
			Next
		Next

		'Return the list of appenders, if any found
		If RetList.Count > 0 Then
			Return RetList
		Else
			Return Nothing
		End If
	End Function

	''' <summary>
	''' Sets the file logging level via an integer value (Overloaded)
	''' </summary>
	''' <param name="InpLevel">Integer corresponding to level (1-5, 5 being most verbose</param>
	''' <remarks></remarks>
	Public Shared Sub SetFileLogLevel(ByVal InpLevel As Integer)

		Dim LogLevelEnumType As Type = GetType(LogLevels)

		'Verify input level is a valid log level
		If Not [Enum].IsDefined(LogLevelEnumType, InpLevel) Then
			WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Invalid value specified for level: " & InpLevel.ToString)
			Return
		End If

		'Convert input integer into the associated enum
		Dim Lvl As LogLevels = DirectCast([Enum].Parse(LogLevelEnumType, InpLevel.ToString), LogLevels)
		SetFileLogLevel(Lvl)

	End Sub

	''' <summary>
	''' Sets file logging level based on enumeration (Overloaded)
	''' </summary>
	''' <param name="InpLevel">LogLevels value defining level (Debug is most verbose)</param>
	''' <remarks></remarks>
	Public Shared Sub SetFileLogLevel(ByVal InpLevel As LogLevels)

		Dim LogRepo As Repository.Hierarchy.Logger = DirectCast(m_FileLogger.Logger, Repository.Hierarchy.Logger)

		Select Case InpLevel
			Case LogLevels.DEBUG
				LogRepo.Level = LogRepo.Hierarchy.LevelMap("DEBUG")
			Case LogLevels.ERROR
				LogRepo.Level = LogRepo.Hierarchy.LevelMap("ERROR")
			Case LogLevels.FATAL
				LogRepo.Level = LogRepo.Hierarchy.LevelMap("FATAL")
			Case LogLevels.INFO
				LogRepo.Level = LogRepo.Hierarchy.LevelMap("INFO")
			Case LogLevels.WARN
				LogRepo.Level = LogRepo.Hierarchy.LevelMap("WARN")
		End Select
	End Sub

	''' <summary>
	''' Creates and configures a file appender
	''' </summary>
	''' <param name="LogfileName">Base of log file to be used</param>
	''' <returns>A configured file appender</returns>
	''' <remarks></remarks>
	Private Shared Function CreateFileAppender(ByVal LogfileName As String) As log4net.Appender.FileAppender
		Dim ReturnAppender As New log4net.Appender.FileAppender()

		ReturnAppender.Name = "FileAppender"
		m_FileDate = DateTime.Now.ToString("MM-dd-yyyy")
		m_BaseFileName = LogfileName
		ReturnAppender.File = (m_BaseFileName & "_") + m_FileDate & ".txt"
		ReturnAppender.AppendToFile = True
		Dim Layout As New log4net.Layout.PatternLayout()
		Layout.ConversionPattern = "%date{MM/dd/yyyy HH:mm:ss}, %message, %level,%newline"
		Layout.ActivateOptions()
		ReturnAppender.Layout = Layout
		ReturnAppender.ActivateOptions()

		Return ReturnAppender
	End Function

	''' <summary>
	''' Configures the file logger
	''' </summary>
	''' <param name="LogFileName">Base name for log file</param>
	''' <param name="LogLevel">Debug level for file logger</param>
	''' <remarks></remarks>
	Public Shared Sub CreateFileLogger(ByVal LogFileName As String, ByVal LogLevel As Integer)
		Dim curLogger As log4net.Repository.Hierarchy.Logger = DirectCast(m_FileLogger.Logger, log4net.Repository.Hierarchy.Logger)
		m_FileAppender = CreateFileAppender(LogFileName)
		curLogger.AddAppender(m_FileAppender)
		SetFileLogLevel(LogLevel)
	End Sub

	''' <summary>
	''' Configures the Db logger
	''' </summary>
	''' <param name="ConnStr">Database connection string</param>
	''' <param name="ModuleName">Module name used by logger</param>
	''' <remarks></remarks>
	Public Shared Sub CreateDbLogger(ByVal ConnStr As String, ByVal ModuleName As String)
		Dim CurLogger As log4net.Repository.Hierarchy.Logger = DirectCast(m_DbLogger.Logger, log4net.Repository.Hierarchy.Logger)
		CurLogger.Level = log4net.Core.Level.Info
		CurLogger.AddAppender(CreateDbAppender(ConnStr, ModuleName))
		CurLogger.AddAppender(m_FileAppender)
	End Sub

	''' <summary>
	''' Creates a database appender
	''' </summary>
	''' <param name="ConnStr">Database connection string</param>
	''' <param name="ModuleName">Module name used by logger</param>
	''' <returns>ADONet database appender</returns>
	''' <remarks></remarks>
	Private Shared Function CreateDbAppender(ByVal ConnStr As String, ByVal ModuleName As String) As log4net.Appender.AdoNetAppender

		Dim ReturnAppender As New log4net.Appender.AdoNetAppender()

		ReturnAppender.Name = "DbAppender"

		ReturnAppender.BufferSize = 1
		ReturnAppender.ConnectionType = "System.Data.SqlClient.SqlConnection, System.Data, Version=1.0.3300.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"
		ReturnAppender.ConnectionString = ConnStr
		ReturnAppender.CommandType = CommandType.StoredProcedure
		ReturnAppender.CommandText = "PostLogEntry"

		'Type parameter
		Dim TypeParam As New log4net.Appender.AdoNetAppenderParameter()
		TypeParam.ParameterName = "@type"
		TypeParam.DbType = DbType.String
		TypeParam.Size = 50
		TypeParam.Layout = CreateLayout("%level")
		ReturnAppender.AddParameter(TypeParam)

		'Message parameter
		Dim MsgParam As New log4net.Appender.AdoNetAppenderParameter()
		MsgParam.ParameterName = "@message"
		MsgParam.DbType = DbType.String
		MsgParam.Size = 4000
		MsgParam.Layout = CreateLayout("%message")
		ReturnAppender.AddParameter(MsgParam)

		'PostedBy parameter
		Dim PostByParam As New log4net.Appender.AdoNetAppenderParameter()
		PostByParam.ParameterName = "@postedBy"
		PostByParam.DbType = DbType.String
		PostByParam.Size = 128
		PostByParam.Layout = CreateLayout(ModuleName)
		ReturnAppender.AddParameter(PostByParam)

		ReturnAppender.ActivateOptions()

		Return ReturnAppender
	End Function

	''' <summary>
	''' Creates a layout object for a Db appender parameter
	''' </summary>
	''' <param name="LayoutStr">Name of parameter</param>
	''' <returns></returns>
	''' <remarks>log4net.Layout.IRawLayout</remarks>
	Private Shared Function CreateLayout(ByVal LayoutStr As String) As log4net.Layout.IRawLayout

		Dim LayoutConvert As New log4net.Layout.RawLayoutConverter()
		Dim ReturnLayout As New log4net.Layout.PatternLayout()
		ReturnLayout.ConversionPattern = LayoutStr
		ReturnLayout.ActivateOptions()
		Return DirectCast(LayoutConvert.ConvertFrom(ReturnLayout), log4net.Layout.IRawLayout)

	End Function
#End Region

End Class
