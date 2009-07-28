'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
' Last modified 07/27/2009
'*********************************************************************************************************

Public Class clsInstData

	'*********************************************************************************************************
	' Class to hold data for each instrument
	'*********************************************************************************************************

#Region "Module variables"
	Private m_StorageVol As String
	Private m_StoragePath As String
	Private m_CapMethod As String
	Private m_InstName As String
#End Region

#Region "Properties"
	Public Property StorageVolume() As String
		Get
			Return m_StorageVol
		End Get
		Set(ByVal value As String)
			m_StorageVol = value
		End Set
	End Property

	Public Property StoragePath() As String
		Get
			Return m_StoragePath
		End Get
		Set(ByVal value As String)
			m_StoragePath = value
		End Set
	End Property

	Public Property CaptureMethod() As String
		Get
			Return m_CapMethod
		End Get
		Set(ByVal value As String)
			m_CapMethod = value
		End Set
	End Property

	Public Property InstName() As String
		Get
			Return m_InstName
		End Get
		Set(ByVal value As String)
			m_InstName = value
		End Set
	End Property
#End Region

End Class
