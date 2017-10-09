'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2009, Battelle Memorial Institute
' Created 08/04/2009
'
' Last modified 08/04/2009
'*********************************************************************************************************
Imports System.Collections.Generic

Public Class clsStatusData

	'*********************************************************************************************************
	'Class to hold long-term data for status reporting. This is a hack to avoid adding an instance of the
	'	status file class to the log tools class
	'*********************************************************************************************************
#Region "Module variables"
    Private Shared m_MostRecentLogMessage As String
    Private Shared ReadOnly m_ErrorQueue As Queue(Of String) = New Queue(Of String)
#End Region

#Region "Properties"
    Public Shared Property MostRecentLogMessage As String
        Get
            Return m_MostRecentLogMessage
        End Get
        Set
            'Filter out routine startup and shutdown messages
            If Value.Contains("=== Started") Or Value.Contains("===== Closing") Then
                'Do nothing
            Else
                m_MostRecentLogMessage = Value
            End If
        End Set
    End Property

    Public Shared ReadOnly Property ErrorQueue As Queue(Of String)
        Get
            Return m_ErrorQueue
        End Get
    End Property
#End Region

#Region "Methods"
    Public Shared Sub AddErrorMessage(ErrMsg As String)
        'Add the most recent error message
        m_ErrorQueue.Enqueue(ErrMsg)

        'If there are > 4 entries in the queue, then delete the oldest ones
        If m_ErrorQueue.Count > 4 Then
            While m_ErrorQueue.Count > 4
                m_ErrorQueue.Dequeue()
            End While
            m_ErrorQueue.TrimExcess()
        End If

    End Sub

#End Region


End Class
