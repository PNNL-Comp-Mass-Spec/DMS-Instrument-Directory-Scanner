'Copyright Pacific Northwest National Laboratory / Battelle
'File:  Main.vb
'File Contents:  Directory scanning methods
'Author(s):  Nathan Trimble
'Comments:  Scans instrument directories and outputs files listing contents of directories.

Option Strict On

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
        Dim swOutFile As System.IO.StreamWriter

        If ds Is Nothing Then
            Return False
        Else
            Dim dt As DataTable = ds.Tables.Item(0)
            Dim drc As DataRowCollection = dt.Rows
            Dim oneRow As DataRow

            For Each oneRow In drc
                instrumentName = CStr(oneRow.Item("Instrument"))
                Dim vol As String = CStr(oneRow.Item("vol"))
                Dim path As String = CStr(oneRow.Item("Path"))

                swOutFile = CreateOutputFile(instrumentName)

                If CStr(oneRow.Item("method")) = "fso" Then
                    GetFSODirectory(instrumentName, vol, path, swOutFile)
                ElseIf CStr(oneRow.Item("method")) = "secfso" Then
                    GetWorkgroupShareDirectory(instrumentName, vol, path, swOutFile)
                ElseIf CStr(oneRow.Item("method")) = "ftp" Then
                    logger.PostEntry("FTP no longer supported, change instrument " & instrumentName & " ""method"" to ""fso"".", ILogger.logMsgType.logError, True)
                End If

                swOutFile.Close()
            Next

            Return True
        End If
    End Function

    Private Function FileSizeToText(ByVal lngFileSizeBytes As Long) As String

        Dim sngFileSize As Single
        Dim intFileSizeIterator As Integer
        Dim strFileSize As String
        Dim strRoundSpec As String

        sngFileSize = CSng(lngFileSizeBytes)

        intFileSizeIterator = 0
        Do While sngFileSize > 1024 And intFileSizeIterator < 3
            sngFileSize /= 1024
            intFileSizeIterator += 1
        Loop

        If sngFileSize < 10 Then
            strRoundSpec = "0.0"
        Else
            strRoundSpec = "0"
        End If

        strFileSize = sngFileSize.ToString(strRoundSpec)

        Select Case intFileSizeIterator
            Case 0
                strFileSize &= " bytes"
            Case 1
                strFileSize &= " KB"
            Case 2
                strFileSize &= " MB"
            Case 3
                strFileSize &= " GB"
            Case Else
                strFileSize &= " ???"
        End Select

        Return strFileSize
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
    Private Sub GetFSODirectory(ByVal instrumentName As String, ByVal serverName As String, ByVal sourcePaths As String, ByRef swOutFile As System.IO.StreamWriter)
        Dim delimiter() As Char = {";"c}
        Dim paths() As String = sourcePaths.Split(delimiter)
        Dim path As String
        Dim strFileSize As String

        For Each path In paths
            path = serverName & path

            Dim dir As New DirectoryInfo(path)
            logger.PostEntry("Reading " & instrumentName & ", Folder " & path, ILogger.logMsgType.logNA, True)
            WriteToOutput(swOutFile, "Folder: " & path)
            If Not dir.Exists Then
                WriteToOutput(swOutFile, "(folder does not exist)")
            Else
                Dim dirs() As DirectoryInfo = dir.GetDirectories
                Dim files() As FileInfo = dir.GetFiles()
                Dim aDir As DirectoryInfo
                Dim aFile As FileInfo

                For Each aDir In dirs
                    WriteToOutput(swOutFile, "Dir", aDir.Name)
                Next
                For Each aFile In files
                    strFileSize = FileSizeToText(aFile.Length)

                    WriteToOutput(swOutFile, "File", aFile.Name, strFileSize)
                Next
            End If
        Next

    End Sub

    'Eventually all servers will be migrated to this function as they are removed from the PNL domain and placed
    'on the BIONET workgroup.  This funtion connects to the shares on a workgroup using ShareConnector.vb
    Private Sub GetWorkgroupShareDirectory(ByVal instrumentName As String, ByVal serverName As String, ByVal sourcePaths As String, ByRef swOutFile As System.IO.StreamWriter)
        Dim delimiter() As Char = {";"c}
        Dim paths() As String = sourcePaths.Split(delimiter)
        Dim path As String
        Dim strFileSize As String

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
            WriteToOutput(swOutFile, "Folder: " & path)
            If Not dir.Exists Then
                WriteToOutput(swOutFile, "(folder does not exist)")
            Else
                Dim dirs() As DirectoryInfo = dir.GetDirectories
                Dim files() As FileInfo = dir.GetFiles()
                Dim aDir As DirectoryInfo
                Dim aFile As FileInfo

                For Each aDir In dirs
                    WriteToOutput(swOutFile, "Dir ", aDir.Name)
                Next
                For Each aFile In files
                    strFileSize = FileSizeToText(aFile.Length)
                    WriteToOutput(swOutFile, "File ", aFile.Name, strFileSize)
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

    Private Function CreateOutputFile(ByVal instrumentName As String) As System.IO.StreamWriter

        Dim outputFilePath As String = Me.outputFileDir & Path.DirectorySeparatorChar & instrumentName & "_source.txt"
        Dim swOutFile As System.IO.StreamWriter

        Try
            System.IO.File.Delete(outputFilePath)
        Catch ex As Exception
            logger.PostEntry("Deletion of file: " & outputFilePath & " could not occur: " & ex.message, ILogger.logMsgType.logError, True)
        End Try

        Try
            swOutFile = New System.IO.StreamWriter(New System.IO.FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))

            ' The file always starts with a blank line (by convention)
            swOutFile.WriteLine()

        Catch ex As Exception
            logger.PostEntry("Creation of file: " & outputFilePath & " failed: " & ex.message, ILogger.logMsgType.logError, True)
        End Try

        Return swOutFile

    End Function

    Private Function WriteToOutput(ByRef swOutFile As System.IO.StreamWriter, _
                                   ByVal field1 As String) As Boolean

        Return WriteToOutput(swOutFile, field1, String.Empty, String.Empty)
    End Function

    Private Function WriteToOutput(ByRef swOutFile As System.IO.StreamWriter, _
                                   ByVal field1 As String, _
                                   ByVal field2 As String) As Boolean

        Return WriteToOutput(swOutFile, field1, field2, String.Empty)
    End Function

    Private Function WriteToOutput(ByRef swOutFile As System.IO.StreamWriter, _
                                   ByVal field1 As String, _
                                   ByVal field2 As String, _
                                   ByVal field3 As String) As Boolean

        Dim strLineOut As String


        strLineOut = field1 & ControlChars.Tab & field2 & ControlChars.Tab & field3
        Debug.WriteLine("WriteToOutput(" & strLineOut & ")")

        swOutFile.WriteLine(strLineOut)

    End Function
End Class
