'*********************************************************************************************************
' Written by Dave Clark and Matthew Monroe for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 01/01/2009
'*********************************************************************************************************

Imports System.IO
Imports System.Text.RegularExpressions
Imports log4net
Imports log4net.Appender

' This assembly attribute tells Log4Net where to find the config file
<Assembly: Config.XmlConfigurator(ConfigFile:="Logging.config", Watch:=True)>

''' <summary>
''' Class for handling logging via Log4Net
''' </summary>
''' <remarks>
''' Call method CreateFileLogger to define the log file name
''' </remarks>
Public NotInheritable Class clsLogTools

#Region "Constants"

    Private Const LOG_FILE_APPENDER = "FileAppender"

    ''' <summary>
    ''' Date format for log file names
    ''' </summary>
    Public Const LOG_FILE_DATECODE = "MM-dd-yyyy"

    Private Const LOG_FILE_MATCH_SPEC = "??-??-????"

    Private Const LOG_FILE_DATE_REGEX = "(?<Month>\d+)-(?<Day>\d+)-(?<Year>\d{4,4})"

    Private Const LOG_FILE_EXTENSION = ".txt"

    Private Const OLD_LOG_FILE_AGE_THRESHOLD_DAYS = 32
#End Region

#Region "Enums"
    ''' <summary>
    ''' Log levels
    ''' </summary>
    Public Enum LogLevels
        ''' <summary>
        ''' Debug message
        ''' </summary>
        DEBUG = 5

        ''' <summary>
        ''' Informational message
        ''' </summary>
        INFO = 4

        ''' <summary>
        ''' Warning message
        ''' </summary>
        WARN = 3

        ''' <summary>
        ''' Error message
        ''' </summary>
        [ERROR] = 2

        ''' <summary>
        ''' Fatal error message
        ''' </summary>
        FATAL = 1
    End Enum

    ''' <summary>
    ''' Log types
    ''' </summary>
    Public Enum LoggerTypes
        ''' <summary>
        ''' Log to a log file
        ''' </summary>
        LogFile

        ''' <summary>
        ''' Log to the database and to the log file
        ''' </summary>
        LogDb

        ''' <summary>
        ''' Log to the system event log and to the log file
        ''' </summary>
        LogSystem
    End Enum
#End Region

#Region "Module variables"
    ''' <summary>
    ''' File Logger (RollingFileAppender)
    ''' </summary>
    Private Shared ReadOnly m_FileLogger As ILog = LogManager.GetLogger("FileLogger")

    ''' <summary>
    ''' Database logger
    ''' </summary>
    Private Shared ReadOnly m_DbLogger As ILog = LogManager.GetLogger("DbLogger")

    ''' <summary>
    ''' System event log logger
    ''' </summary>
    Private Shared ReadOnly m_SysLogger As ILog = LogManager.GetLogger("SysLogger")

    Private Shared m_FileDate As String

    ''' <summary>
    ''' Base log file name
    ''' </summary>
    ''' <remarks>This is updated by ChangeLogFileBaseName or CreateFileLogger</remarks>
    Private Shared m_BaseFileName As String = "UnknownApp"

    Private Shared m_FileAppender As FileAppender

    Private Shared m_LastCheckOldLogs As DateTime = DateTime.UtcNow.AddDays(-1)

    Private Shared m_MostRecentErrorMessage As String = String.Empty

#End Region

#Region "Properties"

    ''' <summary>
    ''' File path for the current log file used by the FileAppender
    ''' </summary>
    ''' <returns>Log file path</returns>
    Public Shared ReadOnly Property CurrentFileAppenderPath As String
        Get
            If (String.IsNullOrEmpty(m_FileAppender?.File)) Then
                Return String.Empty
            End If

            Return m_FileAppender.File
        End Get
    End Property

    ''' <summary>
    ''' Tells calling program file debug status
    ''' </summary>
    ''' <returns>TRUE if debug level enabled for file logger; FALSE otherwise</returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property FileLogDebugEnabled As Boolean
        Get
            Return m_FileLogger.IsDebugEnabled
        End Get
    End Property

    ''' <summary>
    ''' Most recent error message
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property MostRecentErrorMessage As String
        Get
            Return m_MostRecentErrorMessage
        End Get
    End Property

#End Region

#Region "Methods"

    ''' <summary>
    ''' Empty, private constructor to prevent instantiation
    ''' </summary>
    ''' <remarks>
    ''' To simulate a static class in VB.NET, we use the NotInheritable keyword and include a private constructor
    ''' </remarks>
    Private Sub New()

    End Sub

    ''' <summary>
    ''' Write a message to the logging system
    ''' </summary>
    ''' <param name="loggerType">Type of logger to use</param>
    ''' <param name="logLevel">Level of log reporting</param>
    ''' <param name="message">Message to be logged</param>
    Public Shared Sub WriteLog(loggerType As LoggerTypes, logLevel As LogLevels, message As String)
        WriteLogWork(loggerType, logLevel, message, Nothing)
    End Sub

    ''' <summary>
    ''' Write a message and exception to the logging system
    ''' </summary>
    ''' <param name="loggerType">Type of logger to use</param>
    ''' <param name="logLevel">Level of log reporting</param>
    ''' <param name="message">Message to be logged</param>
    ''' <param name="ex">Exception to be logged</param>
    Public Shared Sub WriteLog(loggerType As LoggerTypes, logLevel As LogLevels, message As String, ex As Exception)
        WriteLogWork(loggerType, logLevel, message, ex)
    End Sub

    ''' <summary>
    ''' Write a message and possibly an exception to the logging system
    ''' </summary>
    ''' <param name="loggerType">Type of logger to use</param>
    ''' <param name="logLevel">Level of log reporting</param>
    ''' <param name="message">Message to be logged</param>
    ''' <param name="ex">Exception to be logged</param>
    Public Shared Sub WriteLogWork(loggerType As LoggerTypes, logLevel As LogLevels, message As String, ex As Exception)

        Dim myLogger As ILog

        ' Establish which logger will be used
        Select Case loggerType
            Case LoggerTypes.LogDb
                myLogger = m_DbLogger
                message = Net.Dns.GetHostName() & ": " & message
            Case LoggerTypes.LogFile
                myLogger = m_FileLogger

                ' Check to determine if a new file should be started
                Dim testFileDate As String = DateTime.Now.ToString(LOG_FILE_DATECODE)
                If Not String.Equals(testFileDate, m_FileDate) Then
                    m_FileDate = testFileDate
                    ChangeLogFileName()
                End If
            Case LoggerTypes.LogSystem
                myLogger = m_SysLogger
            Case Else
                Throw New Exception("Invalid logger type specified")
        End Select

        RaiseEvent MessageLogged(message, logLevel)

        If myLogger Is Nothing Then
            Return
        End If

        ' Send the log message
        Select Case logLevel
            Case LogLevels.DEBUG
                If myLogger.IsDebugEnabled Then
                    If ex Is Nothing Then
                        myLogger.Debug(message)
                    Else
                        myLogger.Debug(message, ex)
                    End If
                End If
            Case LogLevels.ERROR
                If myLogger.IsErrorEnabled Then
                    If ex Is Nothing Then
                        myLogger.Error(message)
                    Else
                        myLogger.Error(message, ex)
                    End If
                End If
            Case LogLevels.FATAL
                If myLogger.IsFatalEnabled Then
                    If ex Is Nothing Then
                        myLogger.Fatal(message)
                    Else
                        myLogger.Fatal(message, ex)
                    End If
                End If
            Case LogLevels.INFO
                If myLogger.IsInfoEnabled Then
                    If ex Is Nothing Then
                        myLogger.Info(message)
                    Else
                        myLogger.Info(message, ex)
                    End If
                End If
            Case LogLevels.WARN
                If myLogger.IsWarnEnabled Then
                    If ex Is Nothing Then
                        myLogger.Warn(message)
                    Else
                        myLogger.Warn(message, ex)
                    End If
                End If
            Case Else
                Throw New Exception("Invalid log level specified")
        End Select

        If logLevel <= LogLevels.ERROR Then
            m_MostRecentErrorMessage = message
        End If

        If DateTime.UtcNow.Subtract(m_LastCheckOldLogs).TotalHours > 24 Then
            m_LastCheckOldLogs = DateTime.UtcNow

            Dim curLogger = TryCast(m_FileLogger.Logger, Repository.Hierarchy.Logger)
            If curLogger Is Nothing Then Exit Sub

            For Each item In curLogger.Appenders
                Dim curAppender = TryCast(item, FileAppender)
                If curAppender Is Nothing Then Continue For

                ArchiveOldLogs(curAppender.File)
                Exit For
            Next

        End If
    End Sub

    ''' <summary>
    ''' Update the log file's base name
    ''' </summary>
    ''' <param name="baseName"></param>
    ''' <remarks>Will append today's date to the base name</remarks>
    Public Shared Sub ChangeLogFileBaseName(baseName As String)
        m_BaseFileName = baseName
        ChangeLogFileName()
    End Sub

    ''' <summary>
    '''  Changes the base log file name
    ''' </summary>
    Public Shared Sub ChangeLogFileName()
        m_FileDate = DateTime.Now.ToString(LOG_FILE_DATECODE)
        ChangeLogFileName(m_BaseFileName & "_" & m_FileDate & LOG_FILE_EXTENSION)
    End Sub

    ''' <summary>
    ''' Changes the base log file name
    ''' </summary>
    ''' <param name="relativeFilePath">Log file base name and path (relative to program folder)</param>
    ''' <remarks></remarks>
    Public Shared Sub ChangeLogFileName(relativeFilePath As String)

        ' Get a list of appenders
        Dim appendList As IEnumerable(Of IAppender) = FindAppenders(LOG_FILE_APPENDER)
        If appendList Is Nothing Then
            WriteLog(LoggerTypes.LogSystem, LogLevels.WARN, "Unable to change file name. No appender found")
            Return
        End If

        For Each selectedAppender As IAppender In appendList
            ' Convert the IAppender object to a FileAppender instance
            Dim appenderToChange = TryCast(selectedAppender, FileAppender)
            If appenderToChange Is Nothing Then
                WriteLog(LoggerTypes.LogSystem, LogLevels.ERROR, "Unable to convert appender since not a FileAppender")
                Return
            End If

            ' Change the file name and activate change
            appenderToChange.File = relativeFilePath
            appenderToChange.ActivateOptions()
        Next
    End Sub

    ''' <summary>
    ''' Gets the specified appender
    ''' </summary>
    ''' <param name="appenderName">Name of appender to find</param>
    ''' <returns>List(IAppender) objects if found; NOTHING otherwise</returns>
    Private Shared Function FindAppenders(appenderName As String) As IEnumerable(Of IAppender)

        ' Get a list of the current loggers
        Dim loggerList() As ILog = LogManager.GetCurrentLoggers()
        If loggerList.GetLength(0) < 1 Then
            Return Nothing
        End If

        ' Create a List of appenders matching the criteria for each logger
        Dim retList As New List(Of IAppender)
        For Each testLogger As ILog In loggerList
            For Each testAppender As IAppender In testLogger.Logger.Repository.GetAppenders()
                If testAppender.Name = appenderName Then
                    retList.Add(testAppender)
                End If
            Next
        Next

        ' Return the list of appenders, if any found
        If retList.Count > 0 Then
            Return retList
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Sets the file logging level via an integer value (Overloaded)
    ''' </summary>
    ''' <param name="logLevel">Integer corresponding to level (1-5, 5 being most verbose</param>
    Public Shared Sub SetFileLogLevel(logLevel As Integer)

        Dim logLevelEnumType As Type = GetType(LogLevels)

        ' Verify input level is a valid log level
        If Not [Enum].IsDefined(logLevelEnumType, logLevel) Then
            WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Invalid value specified for level: " & logLevel.ToString)
            Return
        End If

        ' Convert input integer into the associated enum
        Dim logLevelEnum = DirectCast([Enum].Parse(logLevelEnumType, logLevel.ToString), LogLevels)
        SetFileLogLevel(logLevelEnum)

    End Sub

    ''' <summary>
    ''' Sets file logging level based on enumeration (Overloaded)
    ''' </summary>
    ''' <param name="logLevel">LogLevels value defining level (Debug is most verbose)</param>
    Public Shared Sub SetFileLogLevel(logLevel As LogLevels)

        Dim logger = DirectCast(m_FileLogger.Logger, Repository.Hierarchy.Logger)

        Select Case logLevel
            Case LogLevels.DEBUG
                logger.Level = logger.Hierarchy.LevelMap("DEBUG")
            Case LogLevels.ERROR
                logger.Level = logger.Hierarchy.LevelMap("ERROR")
            Case LogLevels.FATAL
                logger.Level = logger.Hierarchy.LevelMap("FATAL")
            Case LogLevels.INFO
                logger.Level = logger.Hierarchy.LevelMap("INFO")
            Case LogLevels.WARN
                logger.Level = logger.Hierarchy.LevelMap("WARN")
        End Select
    End Sub

    ''' <summary>
    ''' Look for log files over 32 days old that can be moved into a subdirectory
    ''' </summary>
    ''' <param name="logFilePath"></param>
    Private Shared Sub ArchiveOldLogs(logFilePath As String)

        Dim targetPath = "??"

        Try
            Dim currentLogFile = New FileInfo(logFilePath)

            Dim matchSpec = "*_" & LOG_FILE_MATCH_SPEC & LOG_FILE_EXTENSION

            Dim logDirectory = currentLogFile.Directory
            If logDirectory Is Nothing Then
                WriteLog(LoggerTypes.LogFile, LogLevels.WARN, "Error archiving old log files; cannot determine the parent directory of " & currentLogFile.FullName)
                Return
            End If

            m_LastCheckOldLogs = DateTime.UtcNow

            Dim logFiles = logDirectory.GetFiles(matchSpec)

            Dim matcher = New Regex(LOG_FILE_DATE_REGEX, RegexOptions.Compiled)

            For Each logFile In logFiles
                Dim match = matcher.Match(logFile.Name)

                If Not match.Success Then
                    Continue For
                End If

                Dim logFileYear = Integer.Parse(match.Groups("Year").Value)
                Dim logFileMonth = Integer.Parse(match.Groups("Month").Value)
                Dim logFileDay = Integer.Parse(match.Groups("Day").Value)

                Dim logDate = New DateTime(logFileYear, logFileMonth, logFileDay)

                If DateTime.Now.Subtract(logDate).TotalDays <= OLD_LOG_FILE_AGE_THRESHOLD_DAYS Then
                    Continue For
                End If

                Dim targetDirectory = New DirectoryInfo(Path.Combine(logDirectory.FullName, logFileYear.ToString()))
                If Not targetDirectory.Exists Then
                    targetDirectory.Create()
                End If

                targetPath = Path.Combine(targetDirectory.FullName, logFile.Name)

                logFile.MoveTo(targetPath)
            Next

        Catch ex As Exception
            WriteLog(LoggerTypes.LogFile, LogLevels.[ERROR], "Error moving old log file to " & targetPath, ex)
        End Try

    End Sub

    ''' <summary>
    ''' Creates a file appender
    ''' </summary>
    ''' <param name="logFileNameBase">Base name for log file</param>
    ''' <returns>A configured file appender</returns>
    Private Shared Function CreateFileAppender(logFileNameBase As String) As FileAppender

        m_FileDate = DateTime.Now.ToString(LOG_FILE_DATECODE)
        m_BaseFileName = logFileNameBase

        Dim layout As New Layout.PatternLayout() With {
            .ConversionPattern = "%date{MM/dd/yyyy HH:mm:ss}, %message, %level,%newline"
        }
        layout.ActivateOptions()

        Dim returnAppender As New FileAppender With {
            .Name = LOG_FILE_APPENDER,
            .File = m_BaseFileName & "_" & m_FileDate & LOG_FILE_EXTENSION,
            .AppendToFile = True,
            .Layout = layout
        }

        returnAppender.ActivateOptions()

        Return returnAppender
    End Function

    ''' <summary>
    ''' Configures the file logger
    ''' </summary>
    ''' <param name="logFileNameBase">Base name for log file</param>
    ''' <param name="logLevel">Debug level for file logger (1-5, 5 being most verbose)</param>
    Public Shared Sub CreateFileLogger(logFileNameBase As String, logLevel As Integer)
        Dim curLogger = DirectCast(m_FileLogger.Logger, Repository.Hierarchy.Logger)
        m_FileAppender = CreateFileAppender(logFileNameBase)
        curLogger.AddAppender(m_FileAppender)

        ArchiveOldLogs(m_FileAppender.File)

        SetFileLogLevel(logLevel)
    End Sub

    ''' <summary>
    ''' Configures the file logger
    ''' </summary>
    ''' <param name="logFileNameBase">Base name for log file</param>
    ''' <param name="logLevel">Debug level for file logger </param>
    Public Shared Sub CreateFileLogger(logFileNameBase As String, logLevel As LogLevels)
        CreateFileLogger(logFileNameBase, CInt(logLevel))
    End Sub

    ''' <summary>
    ''' Configures the database logger
    ''' </summary>
    ''' <param name="connStr">Database connection string</param>
    ''' <param name="moduleName">Module name used by logger</param>
    Public Shared Sub CreateDbLogger(connStr As String, moduleName As String)
        Dim curLogger = DirectCast(m_DbLogger.Logger, Repository.Hierarchy.Logger)
        curLogger.Level = Core.Level.Info
        curLogger.AddAppender(CreateDbAppender(connStr, moduleName, "DbAppender"))

        If (m_FileAppender Is Nothing) Then
            Return
        End If

        Dim addFileAppender = True
        For Each appenderItem In curLogger.Appenders
            If (ReferenceEquals(appenderItem, m_FileAppender)) Then
                addFileAppender = False
                Exit For
            End If
        Next

        If (addFileAppender) Then
            curLogger.AddAppender(m_FileAppender)
        End If

    End Sub

    ''' <summary>
    ''' Creates a database appender
    ''' </summary>
    ''' <param name="connectionString">Database connection string</param>
    ''' <param name="moduleName">Module name used by logger</param>
    ''' <param name="appenderName">Appender name</param>
    ''' <returns>ADONet database appender</returns>
    Private Shared Function CreateDbAppender(connectionString As String, moduleName As String, appenderName As String) As AdoNetAppender

        Dim returnAppender As New AdoNetAppender With {
            .BufferSize = 1,
            .ConnectionType = "System.Data.SqlClient.SqlConnection, System.Data, Version=1.0.3300.0, Culture=neutral, PublicKeyToken=b77a5c561934e089",
            .ConnectionString = connectionString,
            .CommandType = CommandType.StoredProcedure,
            .CommandText = "PostLogEntry",
            .Name = appenderName
        }

        ' Type parameter
        Dim typeParam As New AdoNetAppenderParameter With {
            .ParameterName = "@type",
            .DbType = DbType.String,
            .Size = 50,
            .Layout = CreateLayout("%level")
        }
        returnAppender.AddParameter(typeParam)

        ' Message parameter
        Dim msgParam As New AdoNetAppenderParameter With {
            .ParameterName = "@message",
            .DbType = DbType.String,
            .Size = 4000,
            .Layout = CreateLayout("%message")
        }
        returnAppender.AddParameter(msgParam)

        ' PostedBy parameter
        Dim postByParam As New AdoNetAppenderParameter With {
            .ParameterName = "@postedBy",
            .DbType = DbType.String,
            .Size = 128,
            .Layout = CreateLayout(moduleName)
        }
        returnAppender.AddParameter(postByParam)

        returnAppender.ActivateOptions()

        Return returnAppender
    End Function

    ''' <summary>
    ''' Creates a layout object for a Db appender parameter
    ''' </summary>
    ''' <param name="layoutStr">Name of parameter</param>
    ''' <returns></returns>
    Private Shared Function CreateLayout(layoutStr As String) As Layout.IRawLayout

        Dim layoutConvert As New Layout.RawLayoutConverter()
        Dim returnLayout As New Layout.PatternLayout With {
            .ConversionPattern = layoutStr
        }
        returnLayout.ActivateOptions()

        Dim retItem = DirectCast(layoutConvert.ConvertFrom(returnLayout), Layout.IRawLayout)

        If retItem Is Nothing Then
            Throw New Util.TypeConverters.ConversionNotSupportedException("Error converting a PatternLayout to IRawLayout")
        End If

        Return retItem

    End Function
#End Region

#Region "Events"

    ''' <summary>
    ''' Delegate for event MessageLogged
    ''' </summary>
    Public Delegate Sub MessageLoggedEventHandler(message As String, logLevel As LogLevels)

    ''' <summary>
    ''' This event is raised when a message is logged
    ''' </summary>
    Public Shared Event MessageLogged As MessageLoggedEventHandler
#End Region

End Class
