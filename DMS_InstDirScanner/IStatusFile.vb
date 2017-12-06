'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2006, Battelle Memorial Institute
' Created 06/07/2006
'
'*********************************************************************************************************

''' <summary>
''' Interface used by classes that create and update analysis status file
''' </summary>
Public Interface IStatusFile

#Region "Enums"
    ' Status constants
    Enum EnumMgrStatus As Short
        Stopped
        Stopped_Error
        Running
        Disabled_Local
        Disabled_MC
    End Enum

    Enum EnumTaskStatus As Short
        Stopped
        Requesting
        Running
        Closing
        Failed
        No_Task
    End Enum

    Enum EnumTaskStatusDetail As Short
        Retrieving_Resources
        Running_Tool
        Packaging_Results
        Delivering_Results
        No_Task
    End Enum
#End Region

#Region "Properties"
    Property TaskStartTime As Date

    Property FileNamePath As String

    Property MgrName As String

    Property MgrStatus As EnumMgrStatus

    Property CpuUtilization As Integer

    Property Tool As String

    Property TaskStatus As EnumTaskStatus

    Property Duration As Single

    Property Progress As Single

    Property CurrentOperation As String

    Property TaskStatusDetail As EnumTaskStatusDetail

    Property JobNumber As Integer

    Property JobStep As Integer

    Property Dataset As String

    Property MostRecentJobInfo As String

    Property SpectrumCount As Integer

    Property MessageQueueURI As String

    Property MessageQueueTopic As String

    Property LogToMsgQueue As Boolean
#End Region

#Region "Methods"
    Sub WriteStatusFile()

    Overloads Sub UpdateAndWrite(PercentComplete As Single)

    Overloads Sub UpdateAndWrite(Status As EnumTaskStatusDetail, PercentComplete As Single)

    Overloads Sub UpdateAndWrite(Status As EnumTaskStatusDetail, PercentComplete As Single, DTACount As Integer)

    Sub UpdateStopped(MgrError As Boolean)

    Sub UpdateDisabled(Local As Boolean)

    Sub InitStatusFromFile()

    Sub DisposeMessageQueue()
#End Region

End Interface
