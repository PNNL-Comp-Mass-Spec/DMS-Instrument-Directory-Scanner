'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 07/27/2009
'
'*********************************************************************************************************

''' <summary>
''' Class to hold data for each instrument
''' </summary>
Public Class clsInstData

#Region "Module variables"

#End Region

#Region "Properties"

    Public Property StorageVolume As String

    Public Property StoragePath As String

    Public Property CaptureMethod As String

    Public Property InstName As String

#End Region

    Public Overrides Function ToString() As String
        Return InstName + ": " + StorageVolume + StoragePath
    End Function
End Class
