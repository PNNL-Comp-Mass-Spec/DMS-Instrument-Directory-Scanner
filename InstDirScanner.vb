'Copyright Pacific Northwest National Laboratory / Battelle
'File:  Main.vb
'File Contents:  Directory scanning methods
'Author(s):  Nathan Trimble
'Comments:  Scans instrument directories and ouputs files listing contents of directories.

Imports System.Data
Imports System.IO
Imports DMS_InstDirScanner_NET.Logging

Public Class InstDirScanner
    Private cnStr As String
    Private outputFileDir As String
    Private logger As clsFileLogger

    Private user As String
    Private pwd As String

    Public Sub New(ByVal connection As String, ByVal anOutputFileDir As String, ByRef aLogger As clsFileLogger, ByVal userName As String, ByVal userPwd As String)
        cnStr = connection
        outputFileDir = anOutputFileDir
        logger = aLogger
        user = userName
        pwd = userPwd

        Dim delimiter() As Char = {";"c}
    End Sub

    Public Function PerformInstDirScan(ByVal instrumentName As String) As Boolean
        Dim ds As DataSet = Me.GetInstData(instrumentName)
        If ds Is Nothing Then
            Return False
        Else
            Dim dt As DataTable = ds.Tables.Item(0)
            Dim drc As DataRowCollection = dt.Rows
            Dim oneRow As DataRow

            For Each oneRow In drc
                instrumentName = oneRow.Item("Instrument")
                Dim vol As String = oneRow.Item("vol")
                Dim path As String = oneRow.Item("Path")

                WriteToOutput(instrumentName, "", "", True)
                If oneRow.Item("method") = "fso" Then
                    GetFSODirectory(instrumentName, vol, path)
                ElseIf oneRow.Item("method") = "secfso" Then
                    GetWorkgroupShareDirectory(instrumentName, vol, path)
                ElseIf oneRow.Item("method") = "ftp" Then
                    logger.PostEntry("FTP no longer supported, change instrument " & instrumentName & " ""method"" to ""fso"".", ILogger.logMsgType.logError, True)
                End If
            Next
            Return True
        End If
    End Function

    Private Function GetInstData(ByRef instrumentName As String) As DataSet
        Dim sqlQuery As String
        Dim ds As DataSet
        Dim rows As Integer
        Dim dbTools As New clsDBTools(logger, cnStr)

        sqlQuery = " SELECT t_storage_path.SP_vol_name_server as vol, t_storage_path.SP_path as path," & _
        " T_Instrument_Name.IN_capture_method as method, T_Instrument_Name.IN_name as Instrument" & _
        " FROM T_Instrument_Name INNER JOIN t_storage_path ON" & _
        " T_Instrument_Name.IN_source_path_ID = t_storage_path.SP_path_ID" & _
        " WHERE (T_Instrument_Name.IN_status = 'active')"

        If instrumentName <> "all" Then
            sqlQuery = sqlQuery & " AND (T_Instrument_Name.IN_name = '" & instrumentName & "')"
        End If
        Debug.WriteLine(sqlQuery)

        dbTools.GetDiscDataSet(sqlQuery, ds, rows)
        Return ds
    End Function

    'Function name is historic only.  File System Objects are no longer used in .Net, but the characters "fso"
    'are used in the DMS to indicate the use of SMB/CIFS protocol.
    Private Sub GetFSODirectory(ByVal instrumentName As String, ByVal serverName As String, ByVal sourcePaths As String)
        Dim delimiter() As Char = {";"c}
        Dim paths() As String = sourcePaths.Split(delimiter)
        Dim path As String

        For Each path In paths
            path = serverName & path

            Dim dir As New DirectoryInfo(path)
            logger.PostEntry("Reading " & instrumentName & ", Folder " & path, ILogger.logMsgType.logNA, True)
            WriteToOutput(instrumentName, "Folder: " & path, "")
            If Not dir.Exists Then
                WriteToOutput(instrumentName, "(folder does not exist)", "")
            Else
                Dim dirs() As DirectoryInfo = dir.GetDirectories
                Dim files() As FileInfo = dir.GetFiles()
                Dim aDir As DirectoryInfo
                Dim aFile As FileInfo

                For Each aDir In dirs
                    WriteToOutput(instrumentName, "Dir ", aDir.Name)
                Next
                For Each aFile In files
                    WriteToOutput(instrumentName, "File ", aFile.Name)
                Next
            End If
        Next
    End Sub

    'Eventually all servers will be migrated to this function as they are removed from the PNL domain and placed
    'on the BIONET workgroup.  This funtion connects to the shares on a workgroup using ShareConnector.vb
    Private Sub GetWorkgroupShareDirectory(ByVal instrumentName As String, ByVal serverName As String, ByVal sourcePaths As String)
        Dim delimiter() As Char = {";"c}
        Dim paths() As String = sourcePaths.Split(delimiter)
        Dim path As String

        For Each path In paths
            path = serverName & path
            'Workgroup servers are a special case and need to be handled explicitly.
            'This prevents the client machine from needing a user account with the 
            'same credentials as the user account on the Workgroup server.
            Dim sc As New ShareConnector(path, user, pwd)
            Dim connected As Boolean
            connected = sc.Connect()
            If Not connected Then
                logger.PostEntry("Could not connect to " & path, ILogger.logMsgType.logError, True)
            End If

            Dim dir As New DirectoryInfo(path)
            logger.PostEntry("Reading " & instrumentName & ", Folder " & path, ILogger.logMsgType.logNA, True)
            WriteToOutput(instrumentName, "Folder: " & path, "")
            If Not dir.Exists Then
                WriteToOutput(instrumentName, "(folder does not exist)", "")
            Else
                Dim dirs() As DirectoryInfo = dir.GetDirectories
                Dim files() As FileInfo = dir.GetFiles()
                Dim aDir As DirectoryInfo
                Dim aFile As FileInfo

                For Each aDir In dirs
                    WriteToOutput(instrumentName, "Dir ", aDir.Name)
                Next
                For Each aFile In files
                    WriteToOutput(instrumentName, "File ", aFile.Name)
                Next
            End If

            'More Workgroup specific code...
            If connected Then
                If sc.Disconnect() Then
                    connected = False
                Else
                    logger.PostEntry("Could not dissconnect from " & path, ILogger.logMsgType.logError, True)
                End If
            End If
        Next
    End Sub

    Private Function WriteToOutput(ByVal instrumentName As String, ByVal field1 As String, ByVal field2 As String, Optional ByVal clearFile As Boolean = False) As Boolean
        Dim outputFileName As String = Me.outputFileDir & Path.DirectorySeparatorChar & instrumentName & "_source.txt"
        Dim outputFile As TextWriter

        If clearFile Then
            Try
                File.Delete(outputFileName)
            Catch ex As Exception
                logger.PostEntry("Deletion of file: " & outputFileName & " could not occur.", ILogger.logMsgType.logError, True)
            End Try
            outputFile = File.CreateText(outputFileName)
        Else
            outputFile = File.AppendText(outputFileName)
        End If
        outputFile.WriteLine(field1 & vbTab & field2)
        Debug.WriteLine("WriteToOutput(" & field1 & vbTab & field2 & ")")
        outputFile.Close()
    End Function
End Class
