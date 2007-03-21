'Copyright Pacific Northwest National Laboratory / Battelle
'File:  Main.vb
'File Contents:  Setup and startup.
'Author(s):  Nathan Trimble
'Comments:  Program scans Instrument directories and lists their contents in and output file.  The lists of 
'directories to scan comes from DMS.

Imports System.Data
Imports System.IO
Imports System.Xml.XmlNode
Imports System.Xml.XmlDocument
Imports DMS_InstDirScanner_NET.Logging

Public Module Main

    'Runtime configurations
    Private configFileName As String = "InstDirScan.xml"
    Private dbConnectionString As String
    Private debugging As Boolean = False
    Private outputFilePath As String
    Private outputFileName As String
    Private Const defaultOutputDir As String = "C:\DMS InstSourceDirScans\"
    Private outputDir As String
    Private logFileName As String
    Private logDir As String
    Private instrumentName As String

    Private workgroupUserName
    Private workgroupUserPwd

    Private fileLogger As clsFileLogger

    Public Sub Main()
        GetConfigFileParameters(Path.GetDirectoryName(Application.ExecutablePath) & _
            Path.DirectorySeparatorChar & configFileName)

        If logFileName Is Nothing Then
            logFileName = "InstDirScanner"
        End If
        logDir = Path.GetDirectoryName(Application.ExecutablePath) & Path.DirectorySeparatorChar & logDir
        Dim dirInfo As New DirectoryInfo(logDir)
        If Not dirInfo.Exists() Then
            dirInfo.Create()
        End If
        Debug.WriteLine(logDir)
        fileLogger = New clsFileLogger(logDir & logFileName)

        Try
            fileLogger.PostEntry("===== Start ===== ", ILogger.logMsgType.logNormal, True)
            If GetCLIParameters() Then
                If outputDir Is Nothing Then
                    outputDir = defaultOutputDir
                End If
            Else
                fileLogger.PostEntry("===== Exit, command line arguement error. =====", ILogger.logMsgType.logError, True)
                Exit Sub
            End If

            Dim instDirScanner As InstDirScanner = New InstDirScanner(dbConnectionString, outputDir, fileLogger, _
                workgroupUserName, workgroupUserPwd)

            If Not (instDirScanner.PerformInstDirScan(instrumentName)) Then
                fileLogger.PostEntry("===== Exit, directory scan failed. =====", ILogger.logMsgType.logError, True)
                Exit Sub
            End If

            fileLogger.PostEntry("===== Finish ===== ", ILogger.logMsgType.logNormal, True)
        Catch e As Exception
            fileLogger.PostError("An error occurred see exception message.", e, True)
        End Try

    End Sub

    Private Function GetCLIParameters() As Boolean
        Dim separators As String = ","
        Dim commands As String = Microsoft.VisualBasic.Command()
        Dim args() As String = commands.Split(separators.ToCharArray)

        If UBound(args) > 0 Then
            instrumentName = Trim(args(0))
            outputDir = Trim(args(1))
            Return True
        Else
            fileLogger.PostEntry("Wrong number of arguments", ILogger.logMsgType.logError, True)
            Return False
        End If
    End Function

    Private Sub GetConfigFileParameters(ByVal configFile As String)
        Dim configFileReader As New IniFileReader(configFile, True)
        dbConnectionString = configFileReader.GetIniValue("DatabaseSettings", "ConnectionString")
        logDir = configFileReader.GetIniValue("Output", "LogDir")
        workgroupUserName = configFileReader.GetIniValue("Workgroup", "UserName")
        workgroupUserPwd = PwdDecoder.DecodePassword(configFileReader.GetIniValue("Workgroup", "UserPwd"))
    End Sub
End Module