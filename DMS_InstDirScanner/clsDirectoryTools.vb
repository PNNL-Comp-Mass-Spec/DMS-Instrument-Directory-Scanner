'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 01/01/2009
'
' Last modified 01/01/2009
'*********************************************************************************************************
Imports System.IO
Imports System.Collections.Generic
Imports DMS_InstDirScanner.MgrSettings
Imports DMS_InstDirScanner.clsLogTools

Public Class clsDirectoryTools

	'*********************************************************************************************************
	' Handles all directory access tasks
	'*********************************************************************************************************

#Region "Methods"
	Public Shared Function PerformDirectoryScans(ByVal InstList As List(Of clsInstData), ByVal OutFolder As String, _
	  ByVal MgrSettings As clsMgrSettings, ByVal ProgStatus As IStatusFile) As Boolean

		Dim OutFile As StreamWriter

		Dim Progress As Single
		Dim InstCounter As Integer = 0
		Dim InstCount As Integer = InstList.Count

		ProgStatus.TaskStartTime = Now

		For Each Inst As clsInstData In InstList
			InstCounter += 1
			ProgStatus.Duration = CSng(DateDiff(DateInterval.Hour, ProgStatus.TaskStartTime, Now()))
			Progress = 100 * CSng(InstCounter) / CSng(InstCount)
			ProgStatus.UpdateAndWrite(Progress)
			OutFile = CreateOutputFile(Inst.InstName, OutFolder)
			If OutFile Is Nothing Then Return False

			'Get the directory info an write it
			GetDirectoryData(Inst, OutFile, MgrSettings)

			OutFile.Close()
		Next

		Return True

	End Function

	Private Shared Function CreateOutputFile(ByVal InstName As String, ByVal OutFileDir As String) As StreamWriter

		Dim OutFilePath As String = Path.Combine(OutFileDir, InstName & "_source.txt")
		Dim RetFile As StreamWriter

		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Scanning folder for instrument " & InstName)

		'Check for existing file
		If File.Exists(OutFilePath) Then
			Try
				File.Delete(OutFilePath)
			Catch ex As Exception
				Dim Msg As String = "Exception deleting file " & OutFilePath
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg, ex)
				Return Nothing
			End Try
		End If

		'Create the new file
		Try
			RetFile = New StreamWriter(New FileStream(OutFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))
			'The file always starts with a blank line
			RetFile.WriteLine()
			Return RetFile
		Catch ex As Exception
			Dim Msg As String = "Exception creating output file " & OutFilePath
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg, ex)
			Return Nothing
		End Try

	End Function

	Private Shared Sub GetDirectoryData(ByVal InstData As clsInstData, ByRef OutFile As StreamWriter, ByVal MgrSettings As clsMgrSettings)

		Dim Msg As String
		Dim BionetMachine As Boolean = False
		Dim Connected As Boolean
		Dim InpPath As String = Path.Combine(InstData.StorageVolume, InstData.StoragePath)
		Dim ShareConn As ShareConnector = Nothing

		'If this is a machine on bionet, set up a connection
		If InstData.CaptureMethod.ToLower = "secfso" Then
			BionetMachine = True
			ShareConn = New ShareConnector(InpPath, MgrSettings.GetParam("bionetuser"), DecodePassword(MgrSettings.GetParam("bionetpwd")))
			Connected = ShareConn.Connect()
			If Not Connected Then
				Msg = "Could not connect to " & InpPath
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg)
			End If
		End If

		Dim InpDirInfo As New DirectoryInfo(InpPath)
		Msg = "Reading " & InstData.InstName & ", Folder " & InpPath
		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, Msg)
		WriteToOutput(OutFile, "Folder: " & InpPath)
		If Not Directory.Exists(InpPath) Then
			WriteToOutput(OutFile, "(Folder does not exist)")
		Else
			Dim Dirs() As DirectoryInfo = InpDirInfo.GetDirectories()
			Dim Files() As FileInfo = InpDirInfo.GetFiles()
			For Each TempDir As DirectoryInfo In Dirs
				WriteToOutput(OutFile, "Dir ", TempDir.Name)
			Next
			For Each TempFile As FileInfo In Files
				Dim FileSizeStr As String = FileSizeToText(TempFile.Length)
				WriteToOutput(OutFile, "File ", TempFile.Name, FileSizeStr)
			Next
		End If

		'If this was a bionet machine, disconnect
		If Connected Then
			If ShareConn.Disconnect() Then
				Connected = False
			Else
				Msg = "Could not disconnect from " & InpPath
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "msg")
			End If
		End If

	End Sub

	Private Shared Function FileSizeToText(ByVal InpFileSizeBytes As Long) As String

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

	Private Overloads Shared Function WriteToOutput(ByRef OutFile As StreamWriter, ByVal Field1 As String) As Boolean

		Return WriteToOutput(OutFile, Field1, String.Empty, String.Empty)

	End Function

	Private Overloads Shared Function WriteToOutput(ByRef OutFile As StreamWriter, ByVal Field1 As String, _
			ByVal Field2 As String) As Boolean

		Return WriteToOutput(OutFile, Field1, Field2, String.Empty)

	End Function

	Private Overloads Shared Function WriteToOutput(ByRef OutFile As StreamWriter, ByVal Field1 As String, _
			 ByVal Field2 As String, ByVal Field3 As String) As Boolean

		Dim LineOut As String

		LineOut = Field1 & ControlChars.Tab & Field2 & ControlChars.Tab & Field3
		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, "Write to output (" & LineOut & ")")
		OutFile.WriteLine(LineOut)
		Return True

	End Function

	Private Shared Function DecodePassword(ByVal EnPwd As String) As String
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

#End Region


End Class
