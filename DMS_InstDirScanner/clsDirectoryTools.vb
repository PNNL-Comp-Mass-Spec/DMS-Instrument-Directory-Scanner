﻿'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 01/01/2009
'
'*********************************************************************************************************

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports PRISM.Logging

''' <summary>
''' Handles all directory access tasks
''' </summary>
Public Class clsDirectoryTools
    Inherits PRISM.clsEventNotifier

#Region "Methods"

    Private mDebugLevel As Integer = 1
    Private mMostRecentIOErrorInstrument As String

    Private ReadOnly mFileTools As PRISM.clsFileTools

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        mMostRecentIOErrorInstrument = String.Empty
        mFileTools = New PRISM.clsFileTools()
    End Sub

    Public Function PerformDirectoryScans(
        instList As List(Of clsInstData),
        outFolderPath As String,
        mgrSettings As clsMgrSettings,
        progStatus As IStatusFile) As Boolean

        Dim instCounter = 0
        Dim instCount As Integer = instList.Count

        mDebugLevel = mgrSettings.GetParam("debuglevel", 1)

        progStatus.TaskStartTime = DateTime.UtcNow

        For Each instrument As clsInstData In instList
            Try

                instCounter += 1
                progStatus.Duration = CSng(DateTime.UtcNow.Subtract(progStatus.TaskStartTime).TotalHours())
                Dim progress = 100 * CSng(instCounter) / instCount
                progStatus.UpdateAndWrite(progress)

                OnStatusEvent("Scanning folder for instrument " & instrument.InstName)

                Dim fiSourceFile As FileInfo = Nothing
                Dim swOutFile = CreateOutputFile(instrument.InstName, outFolderPath, fiSourceFile)
                If swOutFile Is Nothing Then Return False

                ' Get the directory info and write it
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
                        OnErrorEvent("Exception copying to MostRecentValid directory", ex)
                    End Try
                End If

            Catch ex As Exception
                LogCriticalError("Error finding files for " & instrument.InstName & " in PerformDirectoryScans: " & ex.Message)
            End Try

        Next

        Return True

    End Function

    Private Function CreateOutputFile(instName As String, outFileDir As String, <Out> ByRef fiStatusFile As FileInfo) As StreamWriter

        fiStatusFile = New FileInfo(Path.Combine(outFileDir, instName & "_source.txt"))

        ' Make a backup copy of the existing file
        If fiStatusFile.Exists Then
            Try
                Dim backupDirectory = New DirectoryInfo(Path.Combine(fiStatusFile.Directory.FullName, "PreviousCopy"))
                If Not backupDirectory.Exists Then backupDirectory.Create()
                fiStatusFile.CopyTo(Path.Combine(backupDirectory.FullName, fiStatusFile.Name), True)
            Catch ex As Exception
                OnErrorEvent("Exception copying " + fiStatusFile.Name + "to PreviousCopy directory", ex)
            End Try

            Dim backupErrorMessage As String = String.Empty
            If Not (mFileTools.DeleteFileWithRetry(fiStatusFile, backupErrorMessage)) Then
                LogErrorToDatabase(backupErrorMessage)
                Return Nothing
            End If

        End If

        ' Create the new file; try up to 3 times

        Dim retriesRemaining = 3
        Dim errorMessage As String = String.Empty

        While retriesRemaining >= 0
            retriesRemaining -= 1

            Try
                Dim swOutFile = New StreamWriter(New FileStream(fiStatusFile.FullName, FileMode.Create, FileAccess.Write, FileShare.Read))

                ' The file always starts with a blank line
                swOutFile.WriteLine()
                Return swOutFile

            Catch ex As Exception
                errorMessage = ex.Message
                ' Delay for 1 second before trying again
                Thread.Sleep(1000)
            End Try

        End While

        OnErrorEvent("Exception creating output file " & fiStatusFile.FullName & ": " & errorMessage)

        Return Nothing

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
      intrumentData As clsInstData,
      swOutFile As TextWriter,
      mgrSettings As IMgrParams) As Boolean

        Dim connected = False
        Dim remoteDirectoryPath As String = Path.Combine(intrumentData.StorageVolume, intrumentData.StoragePath)
        Dim shareConn As PRISM.ShareConnector = Nothing

        Dim userDescription As String

        ' If this is a machine on bionet, set up a connection
        If intrumentData.CaptureMethod.ToLower = "secfso" Then
            Dim bionetUser As String = mgrSettings.GetParam("bionetuser")            ' Typically user ftms (not LCMSOperator)

            If Not bionetUser.Contains("\"c) Then
                ' Prepend this computer's name to the username
                bionetUser = Environment.MachineName & "\" & bionetUser
            End If

            shareConn = New PRISM.ShareConnector(remoteDirectoryPath, bionetUser, DecodePassword(mgrSettings.GetParam("bionetpwd")))
            connected = shareConn.Connect()

            userDescription = " as user " & bionetUser
            If Not connected Then
                OnErrorEvent("Could not connect to " & remoteDirectoryPath & userDescription + "; error code " + shareConn.ErrorMessage)
            ElseIf mDebugLevel >= 5 Then
                OnStatusEvent(" ... connected to " & remoteDirectoryPath & userDescription)
            End If
        Else
            userDescription = " as user " & Environment.UserName
            If remoteDirectoryPath.ToLower().Contains(".bionet") Then
                OnWarningEvent("Warning: Connection to a bionet folder should probably use 'secfso'; currently configured to use 'fso' for " & remoteDirectoryPath)
            End If
        End If

        Dim instrumentDataFolder As New DirectoryInfo(remoteDirectoryPath)

        OnStatusEvent("Reading " & intrumentData.InstName & ", Folder " & remoteDirectoryPath & userDescription)

        ' List the folder path and current date/time on the first line
        ' Will look like this:
        ' (Folder: \\VOrbiETD04.bionet\ProteomicsData\ at 2012-01-23 2:15 PM)
        WriteToOutput(swOutFile, "Folder: " & remoteDirectoryPath & " at " & DateTime.Now().ToString("yyyy-MM-dd hh:mm:ss tt"))

        Dim folderExists = Directory.Exists(remoteDirectoryPath)

        If Not folderExists Then
            WriteToOutput(swOutFile, "(Folder does not exist)")
        Else
            Dim directories = instrumentDataFolder.GetDirectories().ToList()
            Dim files = instrumentDataFolder.GetFiles().ToList()
            For Each datasetDirectory As DirectoryInfo In directories
                Dim totalSizeBytes As Int64 = GetDirectorySize(intrumentData.InstName, datasetDirectory)
                WriteToOutput(swOutFile, "Dir ", datasetDirectory.Name, FileSizeToText(totalSizeBytes))
            Next
            For Each datasetFile As FileInfo In files
                Dim fileSizeText As String = FileSizeToText(datasetFile.Length)
                WriteToOutput(swOutFile, "File ", datasetFile.Name, fileSizeText)
            Next
        End If

        ' If this was a bionet machine, disconnect
        If connected Then
            If Not shareConn.Disconnect() Then
                OnErrorEvent("Could not disconnect from " & remoteDirectoryPath)
            End If
        End If

        Return folderExists

    End Function

    Private Function FileSizeToText(fileSizeBytes As Long) As String

        Dim fileSize = CSng(fileSizeBytes)

        Dim fileSizeIterator = 0
        Do While fileSize > 1024 And fileSizeIterator < 3
            fileSize /= 1024
            fileSizeIterator += 1
        Loop

        Dim formatString As String
        If fileSize < 10 Then
            formatString = "0.0"
        Else
            formatString = "0"
        End If

        Dim fileSizeText = fileSize.ToString(formatString)

        Select Case fileSizeIterator
            Case 0
                fileSizeText &= " bytes"
            Case 1
                fileSizeText &= " KB"
            Case 2
                fileSizeText &= " MB"
            Case 3
                fileSizeText &= " GB"
            Case Else
                fileSizeText &= " ???"
        End Select

        Return fileSizeText

    End Function

    Private Function GetDirectorySize(instrumentName As String, datasetDirectory As DirectoryInfo) As Int64

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

    Private Sub WriteToOutput(swOutFile As TextWriter, field1 As String, Optional ByVal field2 As String = "", Optional ByVal field3 As String = "")

        Dim dataLine = field1 & ControlChars.Tab & field2 & ControlChars.Tab & field3
        swOutFile.WriteLine(dataLine)

    End Sub

    Private Function DecodePassword(encodedPassword As String) As String
        ' Decrypts password received from ini file
        ' Password was created by alternately subtracting or adding 1 to the ASCII value of each character

        Dim tempStr = ""

        Dim indx = 1
        Do While Indx <= Len(encodedPassword)
            Dim charCode = CByte(Asc(Mid(encodedPassword, indx, 1)))
            If indx Mod 2 = 0 Then
                charCode = charCode - CByte(1)
            Else
                charCode = charCode + CByte(1)
            End If
            tempStr = tempStr & Chr(charCode)
            indx = Indx + 1
        Loop

        Return tempStr
    End Function

    ''' <summary>
    ''' Logs an error message to the local log file, unless it is currently between midnight and 1 am
    ''' </summary>
    ''' <param name="errorMessage"></param>
    ''' <remarks></remarks>
    Private Sub LogCriticalError(errorMessage As String)
        If DateTime.Now.Hour = 0 Then
            ' Log this error to the database if it is between 12 am and 1 am
            LogErrorToDatabase(errorMessage)
        Else
            OnErrorEvent(errorMessage)
        End If
    End Sub

    Private Sub LogErrorToDatabase(errorMessage As String)
        LogTools.LogError(errorMessage, Nothing, True)
        OnErrorEvent(errorMessage)
    End Sub

    Private Sub LogIOError(instrumentName As String, errorMessage As String)
        If Not String.IsNullOrEmpty(mMostRecentIOErrorInstrument) AndAlso mMostRecentIOErrorInstrument = instrumentName Then
            Return
        End If

        mMostRecentIOErrorInstrument = String.Copy(instrumentName)

        OnWarningEvent(errorMessage)

    End Sub

#End Region


End Class
