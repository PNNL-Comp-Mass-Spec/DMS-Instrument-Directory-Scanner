'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 01/01/2009
'
'*********************************************************************************************************

Imports System.IO
Imports System.Collections.Generic
Imports DMS_InstDirScanner.MgrSettings
Imports System.Runtime.InteropServices

Public Class clsDirectoryTools

	'*********************************************************************************************************
	' Handles all directory access tasks
    '*********************************************************************************************************

#Region "Methods"

    Protected mDebugLevel As Integer = 1
    Protected mMostRecentIOErrorInstrument As String = String.Empty

    Public Function PerformDirectoryScans(
        ByVal instList As List(Of clsInstData),
        ByVal outFolderPath As String,
        ByVal mgrSettings As clsMgrSettings,
        ByVal progStatus As IStatusFile) As Boolean

        Dim progress As Single
        Dim instCounter As Integer = 0
        Dim instCount As Integer = instList.Count

        mDebugLevel = mgrSettings.GetParam("debuglevel", 1)

        progStatus.TaskStartTime = DateTime.UtcNow

        For Each instrument As clsInstData In instList
            Try

                instCounter += 1
                progStatus.Duration = CSng(DateTime.UtcNow.Subtract(progStatus.TaskStartTime).TotalHours())
                progress = 100 * CSng(instCounter) / CSng(instCount)
                progStatus.UpdateAndWrite(progress)

                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.INFO, "Scanning folder for instrument " & instrument.InstName)

                Dim fiSourceFile As FileInfo = Nothing
                Dim swOutFile = CreateOutputFile(instrument.InstName, outFolderPath, fiSourceFile)
                If swOutFile Is Nothing Then Return False

                'Get the directory info an write it
                Dim folderExists = GetDirectoryData(instrument, swOutFile, mgrSettings)

                swOutFile.Close()

                If folderExists Then
                    ' Copy the file to the MostRecentValid folder
                    Try

                        Dim diTargetDirectory As DirectoryInfo
                        diTargetDirectory = New DirectoryInfo(Path.Combine(outFolderPath, "MostRecentValid"))

                        If Not diTargetDirectory.Exists Then diTargetDirectory.Create()

                        fiSourceFile.CopyTo(Path.Combine(diTargetDirectory.FullName, fiSourceFile.Name), True)

                    Catch ex As Exception
                        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, "Exception copying to MostRecentValid directory", ex)
                    End Try
                End If

            Catch ex As Exception
                LogCriticalError("Error finding files for " & instrument.InstName & " in PerformDirectoryScans: " & ex.Message)
            End Try

        Next

        Return True


    End Function

    Private Function CreateOutputFile(ByVal InstName As String, ByVal OutFileDir As String, <Out> ByRef fiStatusFile As FileInfo) As StreamWriter

        Dim diBackupDirectory As DirectoryInfo
        Dim RetFile As StreamWriter

        fiStatusFile = New FileInfo(Path.Combine(OutFileDir, InstName & "_source.txt"))

        ' Make a backup copy of the existing file
        If fiStatusFile.Exists Then
            Try
                diBackupDirectory = New DirectoryInfo(Path.Combine(fiStatusFile.Directory.FullName, "PreviousCopy"))
                If Not diBackupDirectory.Exists Then diBackupDirectory.Create()
                fiStatusFile.CopyTo(Path.Combine(diBackupDirectory.FullName, fiStatusFile.Name), True)
            Catch ex As Exception
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, "Exception copying to PreviousCopy directory", ex)
            End Try

            Try
                ' Now delete the old copy
                fiStatusFile.Delete()
            Catch ex As Exception
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, "Exception deleting file " & fiStatusFile.FullName, ex)
                Return Nothing
            End Try
        End If

        'Create the new file
        Try
            RetFile = New StreamWriter(New FileStream(fiStatusFile.FullName, FileMode.Create, FileAccess.Write, FileShare.Read))

            'The file always starts with a blank line
            RetFile.WriteLine()
            Return RetFile
        Catch ex As Exception
            clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogDb, clsLogTools.LogLevels.ERROR, "Exception creating output file " & fiStatusFile.FullName, ex)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Query the files and folders on the instrument's shared data folder
    ''' </summary>
    ''' <param name="intrumentData"></param>
    ''' <param name="swOutFile"></param>
    ''' <param name="mgrSettings"></param>
    ''' <returns>True on success, false if the target folder is not found</returns>
    ''' <remarks></remarks>
    Private Function GetDirectoryData(
        ByVal intrumentData As clsInstData,
        ByVal swOutFile As StreamWriter,
        ByVal mgrSettings As clsMgrSettings) As Boolean

        Dim Msg As String
        Dim Connected As Boolean
        Dim InpPath As String = Path.Combine(intrumentData.StorageVolume, intrumentData.StoragePath)
        Dim ShareConn As PRISM.Files.ShareConnector = Nothing

        Dim strUserDescription As String = "as user ??"

        'If this is a machine on bionet, set up a connection
        If intrumentData.CaptureMethod.ToLower = "secfso" Then
            Dim strBionetUser As String = mgrSettings.GetParam("bionetuser")            ' Typically user ftms (not LCMSOperator)

            If Not strBionetUser.Contains("\"c) Then
                ' Prepend this computer's name to the username
                strBionetUser = Environment.MachineName & "\" & strBionetUser
            End If

            ShareConn = New PRISM.Files.ShareConnector(InpPath, strBionetUser, DecodePassword(mgrSettings.GetParam("bionetpwd")))
            Connected = ShareConn.Connect()

            strUserDescription = " as user " & strBionetUser
            If Not Connected Then
                Msg = "Could not connect to " & InpPath & strUserDescription + "; error code " + ShareConn.ErrorMessage
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, Msg)
            ElseIf mDebugLevel >= 5 Then
                Msg = " ... connected to " & InpPath & strUserDescription
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.INFO, Msg)
            End If
        Else
            strUserDescription = " as user " & Environment.UserName
            If InpPath.ToLower().Contains(".bionet") Then
                Msg = "Warning: Connection to a bionet folder should probably use 'secfso'; currently configured to use 'fso' for " & InpPath
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.WARN, Msg)
            End If
        End If

        Dim diInstDataFolder As New DirectoryInfo(InpPath)

        Msg = "Reading " & intrumentData.InstName & ", Folder " & InpPath & strUserDescription
        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.DEBUG, Msg)

        ' List the folder path and current date/time on the first line
        ' Will look like this:
        ' (Folder: \\VOrbiETD04.bionet\ProteomicsData\ at 2012-01-23 2:15 PM)
        WriteToOutput(swOutFile, "Folder: " & InpPath & " at " & DateTime.Now().ToString("yyyy-MM-dd hh:mm:ss tt"))

        Dim folderExists = Directory.Exists(InpPath)

        If Not folderExists Then
            WriteToOutput(swOutFile, "(Folder does not exist)")
        Else
            Dim directories = diInstDataFolder.GetDirectories().ToList()
            Dim files = diInstDataFolder.GetFiles().ToList()
            For Each datasetDirectory As DirectoryInfo In directories
                Dim totalSizeBytes As Int64 = GetDirectorySize(intrumentData.InstName, datasetDirectory)
                WriteToOutput(swOutFile, "Dir ", datasetDirectory.Name, FileSizeToText(totalSizeBytes))
            Next
            For Each datasetFile As FileInfo In files
                Dim FileSizeStr As String = FileSizeToText(datasetFile.Length)
                WriteToOutput(swOutFile, "File ", datasetFile.Name, FileSizeStr)
            Next
        End If

        'If this was a bionet machine, disconnect
        If Connected Then
            If Not ShareConn.Disconnect() Then
                Msg = "Could not disconnect from " & InpPath
                clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, Msg)
            End If
        End If

        Return folderExists

    End Function

    Private Function FileSizeToText(ByVal InpFileSizeBytes As Long) As String

        Dim FileSize As Single
        Dim FileSizeIterator As Integer
        Dim FileSizeStr As String
        Dim RoundSpec As String

        FileSize = CSng(InpFileSizeBytes)

        FileSizeIterator = 0
        Do While FileSize > 1024 And FileSizeIterator < 3
            FileSize /= 1024
            FileSizeIterator += 1
        Loop

        If FileSize < 10 Then
            RoundSpec = "0.0"
        Else
            RoundSpec = "0"
        End If

        FileSizeStr = FileSize.ToString(RoundSpec)

        Select Case FileSizeIterator
            Case 0
                FileSizeStr &= " bytes"
            Case 1
                FileSizeStr &= " KB"
            Case 2
                FileSizeStr &= " MB"
            Case 3
                FileSizeStr &= " GB"
            Case Else
                FileSizeStr &= " ???"
        End Select

        Return FileSizeStr

    End Function

    Protected Function GetDirectorySize(ByVal instrumentName As String, ByVal datasetDirectory As DirectoryInfo) As Int64

        Dim totalSizeBytes As Int64 = 0

        Try
            Dim files = datasetDirectory.GetFiles("*").ToList()

            For Each currentFile In files
                totalSizeBytes += currentFile.Length
            Next

            ' Do not use SearchOption.AllDirectories in case there are security issues with one or more of the subfolders
            Dim folders = datasetDirectory.GetDirectories("*").ToList()
            For Each subDirectory In folders
                ' Recursively call this function
                totalSizeBytes += GetDirectorySize(instrumentName, subDirectory)
            Next

        Catch ex As UnauthorizedAccessException
            ' Ignore this error
        Catch ex As IOException
            LogIOError(instrumentName, "IOException determining directory size for " & instrumentName & ", folder " & datasetDirectory.Name & "; ignoring additional errors for this instrument")
        Catch ex As Exception
            LogCriticalError("Error determining directory size for " & instrumentName & ", folder " & datasetDirectory.Name & ": " & ex.Message)
        End Try

        Return totalSizeBytes

    End Function


    Private Sub WriteToOutput(ByVal swOutFile As StreamWriter, ByVal field1 As String, Optional ByVal field2 As String = "", Optional ByVal field3 As String = "")

        Dim LineOut As String

        LineOut = field1 & ControlChars.Tab & field2 & ControlChars.Tab & field3
        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.DEBUG, "Write to output (" & LineOut & ")")
        swOutFile.WriteLine(LineOut)

    End Sub

    Private Function DecodePassword(ByVal EnPwd As String) As String
        'Decrypts password received from ini file
        ' Password was created by alternately subtracting or adding 1 to the ASCII value of each character

        Dim CharCode As Byte
        Dim TempStr As String
        Dim Indx As Integer

        TempStr = ""

        Indx = 1
        Do While Indx <= Len(EnPwd)
            CharCode = CByte(Asc(Mid(EnPwd, Indx, 1)))
            If Indx Mod 2 = 0 Then
                CharCode = CharCode - CByte(1)
            Else
                CharCode = CharCode + CByte(1)
            End If
            TempStr = TempStr & Chr(CharCode)
            Indx = Indx + 1
        Loop

        Return TempStr
    End Function

    ''' <summary>
    ''' Logs an error message to the local log file, unless it is currently between midnight and 1 am
    ''' </summary>
    ''' <param name="errorMessage"></param>
    ''' <remarks></remarks>
    Private Sub LogCriticalError(ByVal errorMessage As String)
        If DateTime.Now.Hour = 0 Then
            ' Log this error to the database if it is between 12 am and 1 am
            clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogDb, clsLogTools.LogLevels.ERROR, errorMessage)
        Else
            clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, errorMessage)
        End If
    End Sub

    Private Sub LogIOError(ByVal instrumentName As String, ByVal errorMessage As String)
        If Not String.IsNullOrEmpty(mMostRecentIOErrorInstrument) AndAlso mMostRecentIOErrorInstrument = instrumentName Then
            Return
        End If

        mMostRecentIOErrorInstrument = String.Copy(instrumentName)

        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.WARN, errorMessage)

    End Sub

#End Region


End Class
