'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2006, Battelle Memorial Institute
' Created 06/07/2006
'
' Last modified 01/16/2008
'						- 05/20/2009 (DAC) - Modified for use of new status file format
'*********************************************************************************************************

Public Interface IStatusFile

	'*********************************************************************************************************
	'Interface used by classes that create and update analysis status file
	'*********************************************************************************************************

#Region "Enums"
	'Status constants
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
	ReadOnly Property StartTime() As Date

	Property FileNamePath() As String

	Property MgrName() As String

	Property MgrStatus() As EnumMgrStatus

	Property CpuUtilization() As Integer

	Property Tool() As String

	Property TaskStatus() As EnumTaskStatus

	Property Duration() As Single

	Property Progress() As Single

	Property CurrentOperation() As String

	Property TaskStatusDetail() As EnumTaskStatusDetail

	Property JobNumber() As Integer

	Property JobStep() As Integer

	Property Dataset() As String

	Property MostRecentJobInfo() As String

	Property SpectrumCount() As Integer

	Property MessageQueueURI() As String

	Property MessageQueueTopic() As String

	Property LogToMsgQueue() As Boolean
#End Region

#Region "Methods"
	Sub WriteStatusFile()

	Overloads Sub UpdateAndWrite(ByVal PercentComplete As Single)

	Overloads Sub UpdateAndWrite(ByVal Status As EnumTaskStatusDetail, ByVal PercentComplete As Single)

	Overloads Sub UpdateAndWrite(ByVal Status As EnumTaskStatusDetail, ByVal PercentComplete As Single, ByVal DTACount As Integer)

	Sub UpdateStopped(ByVal MgrError As Boolean)

	Sub UpdateDisabled(ByVal Local As Boolean)

#End Region

End Interface
