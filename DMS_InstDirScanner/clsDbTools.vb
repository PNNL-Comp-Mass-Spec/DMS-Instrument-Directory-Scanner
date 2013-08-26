'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
' Last modified 07/27/2009
'*********************************************************************************************************

Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Data.SqlClient
Imports DMS_InstDirScanner.MgrSettings
Imports DMS_InstDirScanner.clsLogTools
Imports System.Data

Public Class clsDbTools

	'*********************************************************************************************************
	' Class for handling database access
	'*********************************************************************************************************

#Region "Module variables"
	Private Shared m_ErrorList As New StringCollection()
#End Region

#Region "Methods"
	''' <summary>
	''' Gets a list of instruments and data paths from DMS
	''' </summary>
	''' <param name="MgrSettings">Manager params object</param>
	''' <returns>List(Of clsInstData) containing data for all active instruments</returns>
	''' <remarks></remarks>
	Public Shared Function GetInstrumentList(ByVal MgrSettings As clsMgrSettings) As List(Of clsInstData)

		clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, "Getting instrument list")

		Dim sqlQuery As String = "SELECT * FROM V_Instrument_Source_Paths"

		' Get a table containing the active instruments
		Dim Dt As DataTable = GetDataTable(sqlQuery, MgrSettings.GetParam("connectionstring"))

		' Verify valid data found
		If Dt Is Nothing Then
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "Unable to retrieve instrument list")
			Return Nothing
		End If
		If Dt.Rows.Count < 1 Then
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, "No instruments found")
			Return Nothing
		End If

		' Create a list of all instrument data
		Dim ReturnList As New List(Of clsInstData)
		Dim CurRow As DataRow
		Try
			For Each CurRow In Dt.Rows
				Dim TempData As New clsInstData
				TempData.CaptureMethod = CStr(CurRow(Dt.Columns("method")))
				TempData.InstName = CStr(CurRow(Dt.Columns("Instrument")))
				TempData.StoragePath = CStr(CurRow(Dt.Columns("path")))
				TempData.StorageVolume = CStr(CurRow(Dt.Columns("vol")))
				ReturnList.Add(TempData)
			Next
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.DEBUG, "Retrieved instrument list")
			Return ReturnList
		Catch ex As Exception
			Dim Msg As String = "Exception filling instrument list"
			clsLogTools.WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg, ex)
			Return Nothing
		End Try

	End Function

	Private Shared Function GetDataTable(ByVal SqlStr As String, ByVal ConnStr As String) As DataTable

		Dim Msg As String

		Dim Dt As DataTable = Nothing

		Try
			Using Cn As SqlConnection = New SqlConnection(ConnStr)
				AddHandler Cn.InfoMessage, New SqlInfoMessageEventHandler(AddressOf OnInfoMessage)
				Using Da As SqlDataAdapter = New SqlDataAdapter(SqlStr, Cn)
					Using Ds As DataSet = New DataSet
						Da.Fill(Ds)
						If Ds.Tables.Count = 1 Then
							Dt = Ds.Tables(0)
						Else
							Msg = "Invalid table count: " & Ds.Tables.Count.ToString() & ", while getting instrument list"
							WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg)
							LogErrorEvents()
							Return Nothing
						End If
					End Using 'Ds
				End Using 'Da
				RemoveHandler Cn.InfoMessage, AddressOf OnInfoMessage
			End Using 'Cn
			LogErrorEvents()
			Return Dt
		Catch ex As Exception
			Msg = "Exception while getting list of active instruments"
			WriteLog(LoggerTypes.LogFile, LogLevels.ERROR, Msg, ex)
			Return Nothing
		End Try

	End Function

	''' <summary>
	''' Reports database errors to local log
	''' </summary>
	''' <remarks></remarks>
	Protected Shared Sub LogErrorEvents()
		If m_ErrorList.Count > 0 Then
			WriteLog(LoggerTypes.LogFile, LogLevels.WARN, "Warning messages were posted to local log")
		End If
		Dim s As String
		For Each s In m_ErrorList
			WriteLog(LoggerTypes.LogFile, LogLevels.WARN, s)
		Next
	End Sub

	''' <summary>
	''' Event handler for InfoMessage event
	''' </summary>
	''' <param name="sender"></param>
	''' <param name="args"></param>
	''' <remarks>Errors and warnings from SQL Server are caught here</remarks>
	Private Shared Sub OnInfoMessage(ByVal sender As Object, ByVal args As SqlInfoMessageEventArgs)

		Dim err As SqlError
		Dim s As String
		For Each err In args.Errors
			s = ""
			s &= "Message: " & err.Message
			s &= ", Source: " & err.Source
			s &= ", Class: " & err.Class
			s &= ", State: " & err.State
			s &= ", Number: " & err.Number
			s &= ", LineNumber: " & err.LineNumber
			s &= ", Procedure:" & err.Procedure
			s &= ", Server: " & err.Server
			m_ErrorList.Add(s)
		Next

	End Sub
#End Region

End Class
