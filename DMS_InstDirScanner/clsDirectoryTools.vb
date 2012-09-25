'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 01/01/2009
'
' Last modified 10/24/2011
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
		Dim FolderMissing As Boolean
		Dim fiSourceFile As System.IO.FileInfo

		ProgStatus.TaskStartTime = System.DateTime.UtcNow

		For Each Inst As clsInstData In InstList
			InstCounter += 1
			ProgStatus.Duration = CSng(System.DateTime.UtcNow.Subtract(ProgStatus.TaskStartTime).TotalHours())
			Progress = 100 * CSng(InstCounter) / CSng(InstCount)
			ProgStatus.UpdateAndWrite(Progress)

			fiSourceFile = Nothing
			OutFile = CreateOutputFile(Inst.InstName, OutFolder, fiSourceFile)
			If OutFile Is Nothing Then Return False

			'Get the directory info an write it
			FolderMissing = False
			GetDirectoryData(Inst, OutFile, MgrSettings, FolderMissing)

			OutFile.Close()

			If Not FolderMissing Then
				' Copy the file to the MostRecentValid folder
				Try

					Dim diTargetDirectory As System.IO.DirectoryInfo
					diTargetDirectory = New System.IO.DirectoryInfo(System.IO.Path.Combine(OutFolder, "MostRecentValid"))

					If Not diTargetDirectory.Exists Then diTargetDirectory.Create()

					fiSourceFile.CopyTo(System.IO.Path.Combine(diTargetDirectory.FullName, fiSourceFile.Name), True)

				Catch ex As Exception
					clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Exception copying to MostRecentValid directory", ex)
				End Try
			End If
		Next

		Return True

	End Function

	Private Shared Function CreateOutputFile(ByVal InstName As String, ByVal OutFileDir As String, ByRef fiStatusFile As System.IO.FileInfo) As StreamWriter

		Dim diBackupDirectory As System.IO.DirectoryInfo
		Dim RetFile As StreamWriter

		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.INFO, "Scanning folder for instrument " & InstName)
		fiStatusFile = New System.IO.FileInfo(System.IO.Path.Combine(OutFileDir, InstName & "_source.txt"))

		' Make a backup copy of the existing file
		If fiStatusFile.Exists Then
			Try
				diBackupDirectory = New System.IO.DirectoryInfo(System.IO.Path.Combine(fiStatusFile.Directory.FullName, "PreviousCopy"))
				If Not diBackupDirectory.Exists Then diBackupDirectory.Create()
				fiStatusFile.CopyTo(System.IO.Path.Combine(diBackupDirectory.FullName, fiStatusFile.Name), True)
			Catch ex As Exception
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Exception copying to PreviousCopy directory", ex)
			End Try

			Try
				' Now delete the old copy
				fiStatusFile.Delete()
			Catch ex As Exception
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Exception deleting file " & fiStatusFile.FullName, ex)
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
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Exception creating output file " & fiStatusFile.FullName, ex)
			Return Nothing
		End Try

	End Function

	Private Shared Sub GetDirectoryData(ByVal InstData As clsInstData, ByRef OutFile As StreamWriter, ByVal MgrSettings As clsMgrSettings, ByRef FolderMissing As Boolean)

		Dim Msg As String
		Dim BionetMachine As Boolean = False
		Dim Connected As Boolean
		Dim InpPath As String = Path.Combine(InstData.StorageVolume, InstData.StoragePath)
		Dim ShareConn As ShareConnector = Nothing

		Dim strUserDescription As String = "as user ??"
		FolderMissing = False

		'If this is a machine on bionet, set up a connection
		If InstData.CaptureMethod.ToLower = "secfso" Then
			BionetMachine = True
			Dim strBionetUser As String = MgrSettings.GetParam("bionetuser")			' Typically user ftms (not LCMSOperator)
			ShareConn = New ShareConnector(InpPath, strBionetUser, DecodePassword(MgrSettings.GetParam("bionetpwd")))
			Connected = ShareConn.Connect()

			strUserDescription = " as user " & strBionetUser
			If Not Connected Then
				Msg = "Could not connect to " & InpPath & strUserDescription
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg)
			End If
		Else
			strUserDescription = " as user " & Environment.UserName
			If InpPath.ToLower().Contains(".bionet") Then
				Msg = "Warning: Connection to a bionet folder should probably use 'secfso'; currently configured to use 'fso'"
				clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.WARN, Msg)
			End If
		End If

		Dim diInstDataFolder As New DirectoryInfo(InpPath)

		Msg = "Reading " & InstData.InstName & ", Folder " & InpPath & strUserDescription
		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, Msg)

		' List the folder path and current date/time on the first line
		' Will look like this:
		' (Folder: \\VOrbiETD04.bionet\ProteomicsData\ at 2012-01-23 2:15 PM)
		WriteToOutput(OutFile, "Folder: " & InpPath & " at " & System.DateTime.Now().ToString("yyyy-MM-dd hh:mm:ss tt"))

		If Not Directory.Exists(InpPath) Then
			WriteToOutput(OutFile, "(Folder does not exist)")
			FolderMissing = True
		Else
			Dim Dirs() As DirectoryInfo = diInstDataFolder.GetDirectories()
			Dim Files() As FileInfo = diInstDataFolder.GetFiles()
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
