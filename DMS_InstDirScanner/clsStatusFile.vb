'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy 
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2006, Battelle Memorial Institute
' Created 06/07/2006
'
' Last modified 01/16/2008
'						- 05/20/2009 (DAC) - Modified for use of new status file format
'						- 08/11/2009 (DAC) - Added additional properties and methods for status reporting
'						- 08/20/2009 (DAC) - Added duration in minutes to output
'*********************************************************************************************************

Imports System.IO
Imports System.Xml
Imports MessageLogger
Imports System.Collections.Generic

Public Class clsStatusFile
	Implements IStatusFile

	'*********************************************************************************************************
	'Provides tools for creating and updating an analysis status file
	'*********************************************************************************************************

#Region "Module variables"
	'Status file name and location
	Private m_FileNamePath As String = ""

	'Manager name
	Private m_MgrName As String = ""

	'Status value
	Private m_MgrStatus As IStatusFile.EnumMgrStatus = IStatusFile.EnumMgrStatus.STOPPED

	'Manager start time
	Private m_MgrStartTime As Date

	'Task start time
	Private m_TaskStartTime As Date

	'CPU utilization
	Private m_CpuUtilization As Integer = 0

	'Analysis Tool
	Private m_Tool As String = ""

	'Task status
	Private m_TaskStatus As IStatusFile.EnumTaskStatus = IStatusFile.EnumTaskStatus.NO_TASK

	'Task duration
	Private m_Duration As Single = 0

	'Progess (in percent)
	Private m_Progress As Single = 0

	'Current operation (freeform string)
	Private m_CurrentOperation As String = ""

	'Task status detail
	Private m_TaskStatusDetail As IStatusFile.EnumTaskStatusDetail = IStatusFile.EnumTaskStatusDetail.NO_TASK

	'Job number
	Private m_JobNumber As Integer = 0

	'Job step
	Private m_JobStep As Integer = 0

	'Dataset name
	Private m_Dataset As String = ""

	'Most recent job info
	Private m_MostRecentJobInfo As String = ""

	'Number of spectrum files created
	Private m_SpectrumCount As Integer = 0

	'Message broker connection string
	Private m_MessageQueueURI As String

	'Broker topic for status reporting
	Private m_MessageQueueTopic As String

	'Flag to indicate if status should be logged to broker in addition to a file
	Private m_LogToMsgQueue As Boolean
#End Region

#Region "Properties"
	Public Property FileNamePath() As String Implements IStatusFile.FileNamePath
		Get
			Return m_FileNamePath
		End Get
		Set(ByVal value As String)
			m_FileNamePath = value
		End Set
	End Property

	Public Property MgrName() As String Implements IStatusFile.MgrName
		Get
			Return m_MgrName
		End Get
		Set(ByVal Value As String)
			m_MgrName = Value
		End Set
	End Property

	Public Property MgrStatus() As IStatusFile.EnumMgrStatus Implements IStatusFile.MgrStatus
		Get
			Return m_MgrStatus
		End Get
		Set(ByVal Value As IStatusFile.EnumMgrStatus)
			m_MgrStatus = Value
		End Set
	End Property

	Public Property CpuUtilization() As Integer Implements IStatusFile.CpuUtilization
		Get
			Return m_CpuUtilization
		End Get
		Set(ByVal value As Integer)
			m_CpuUtilization = value
		End Set
	End Property

	Public Property Tool() As String Implements IStatusFile.Tool
		Get
			Return m_Tool
		End Get
		Set(ByVal Value As String)
			m_Tool = Value
		End Set
	End Property

	Public Property TaskStartTime() As Date Implements IStatusFile.TaskStartTime
		Get
			Return m_TaskStartTime
		End Get
		Set(ByVal value As Date)
			m_TaskStartTime = value
		End Set
	End Property

	Public Property TaskStatus() As IStatusFile.EnumTaskStatus Implements IStatusFile.TaskStatus
		Get
			Return m_TaskStatus
		End Get
		Set(ByVal value As IStatusFile.EnumTaskStatus)
			m_TaskStatus = value
		End Set
	End Property

	Public Property Duration() As Single Implements IStatusFile.Duration
		Get
			Return m_Duration
		End Get
		Set(ByVal value As Single)
			m_Duration = value
		End Set
	End Property

	Public Property Progress() As Single Implements IStatusFile.Progress
		Get
			Return m_Progress
		End Get
		Set(ByVal Value As Single)
			m_Progress = Value
		End Set
	End Property

	Public Property CurrentOperation() As String Implements IStatusFile.CurrentOperation
		Get
			Return m_CurrentOperation
		End Get
		Set(ByVal value As String)
			m_CurrentOperation = value
		End Set
	End Property

	Public Property TaskStatusDetail() As IStatusFile.EnumTaskStatusDetail Implements IStatusFile.TaskStatusDetail
		Get
			Return m_TaskStatusDetail
		End Get
		Set(ByVal value As IStatusFile.EnumTaskStatusDetail)
			m_TaskStatusDetail = value
		End Set
	End Property

	Public Property JobNumber() As Integer Implements IStatusFile.JobNumber
		Get
			Return m_JobNumber
		End Get
		Set(ByVal Value As Integer)
			m_JobNumber = Value
		End Set
	End Property

	Public Property JobStep() As Integer Implements IStatusFile.JobStep
		Get
			Return m_JobStep
		End Get
		Set(ByVal value As Integer)
			m_JobStep = value
		End Set
	End Property

	Public Property Dataset() As String Implements IStatusFile.Dataset
		Get
			Return m_Dataset
		End Get
		Set(ByVal Value As String)
			m_Dataset = Value
		End Set
	End Property

	Public Property MostRecentJobInfo() As String Implements IStatusFile.MostRecentJobInfo
		Get
			Return m_MostRecentJobInfo
		End Get
		Set(ByVal value As String)
			m_MostRecentJobInfo = value
		End Set
	End Property

	Public Property SpectrumCount() As Integer Implements IStatusFile.SpectrumCount
		Get
			Return m_SpectrumCount
		End Get
		Set(ByVal value As Integer)
			m_SpectrumCount = value
		End Set
	End Property

	Public Property MessageQueueURI() As String Implements IStatusFile.MessageQueueURI
		Get
			Return m_MessageQueueURI
		End Get
		Set(ByVal value As String)
			m_MessageQueueURI = value
		End Set
	End Property

	Public Property MessageQueueTopic() As String Implements IStatusFile.MessageQueueTopic
		Get
			Return m_MessageQueueTopic
		End Get
		Set(ByVal value As String)
			m_MessageQueueTopic = value
		End Set
	End Property

	Public Property LogToMsgQueue() As Boolean Implements IStatusFile.LogToMsgQueue
		Get
			Return m_LogToMsgQueue
		End Get
		Set(ByVal value As Boolean)
			m_LogToMsgQueue = value
		End Set
	End Property
#End Region

#Region "Methods"
	''' <summary>
	''' Constructor
	''' </summary>
	''' <param name="FileLocation">Full path to status file</param>
	''' <remarks></remarks>
	Public Sub New(ByVal FileLocation As String)
		m_FileNamePath = FileLocation
		m_MgrStartTime = Now()
		m_Progress = 0
		m_SpectrumCount = 0
		m_Dataset = ""
		m_JobNumber = 0
		m_Tool = ""
	End Sub

	''' <summary>
	''' Converts the manager status enum to a string value
	''' </summary>
	''' <param name="StatusEnum">An IStatusFile.EnumMgrStatus object</param>
	''' <returns>String representation of input object</returns>
	''' <remarks></remarks>
	Private Function ConvertMgrStatusToString(ByVal StatusEnum As IStatusFile.EnumMgrStatus) As String

		Return StatusEnum.ToString("G")

	End Function

	''' <summary>
	''' Converts the task status enum to a string value
	''' </summary>
	''' <param name="StatusEnum">An IStatusFile.EnumTaskStatus object</param>
	''' <returns>String representation of input object</returns>
	''' <remarks></remarks>
	Private Function ConvertTaskStatusToString(ByVal StatusEnum As IStatusFile.EnumTaskStatus) As String

		Return StatusEnum.ToString("G")

	End Function

	''' <summary>
	''' Converts the manager status enum to a string value
	''' </summary>
	''' <param name="StatusEnum">An IStatusFile.EnumTaskStatusDetail object</param>
	''' <returns></returns>
	''' <remarks>String representation of input object</remarks>
	Private Function ConvertTaskDetailStatusToString(ByVal StatusEnum As IStatusFile.EnumTaskStatusDetail) As String

		Return StatusEnum.ToString("G")

	End Function

	''' <summary>
	''' Writes the status file
	''' </summary>
	''' <remarks></remarks>
	Public Sub WriteStatusFile() Implements IStatusFile.WriteStatusFile

		'Writes a status file for external monitor to read

		Dim XDocument As System.Xml.XmlDocument
		Dim XWriter As XmlTextWriter

		Dim MemStream As MemoryStream
		Dim MemStreamReader As StreamReader

		Dim XMLText As String = String.Empty

		'Set up the XML writer
		Try
			XDocument = New System.Xml.XmlDocument

			'Create a memory stream to write the document in
			MemStream = New MemoryStream
			XWriter = New XmlTextWriter(MemStream, System.Text.Encoding.UTF8)

			XWriter.Formatting = Formatting.Indented
			XWriter.Indentation = 2

			'Write the file
			XWriter.WriteStartDocument(True)
			'Root level element
			XWriter.WriteStartElement("Root")
			XWriter.WriteStartElement("Manager")
			XWriter.WriteElementString("MgrName", m_MgrName)
			XWriter.WriteElementString("MgrStatus", ConvertMgrStatusToString(m_MgrStatus))
			XWriter.WriteElementString("LastUpdate", Now().ToString)
			XWriter.WriteElementString("LastStartTime", m_MgrStartTime.ToString())
			XWriter.WriteElementString("CPUUtilization", m_CpuUtilization.ToString())
			XWriter.WriteElementString("FreeMemoryMB", "0")
			'TODO: Figure out how to retrieve recent error messages
			XWriter.WriteStartElement("RecentErrorMessages")
			For Each ErrMsg As String In clsStatusData.ErrorQueue
				XWriter.WriteElementString("ErrMsg", ErrMsg)
			Next
			XWriter.WriteEndElement()	'Error messages
			XWriter.WriteEndElement()	'Manager section

			XWriter.WriteStartElement("Task")
			XWriter.WriteElementString("Tool", m_Tool)
			XWriter.WriteElementString("Status", ConvertTaskStatusToString(m_TaskStatus))
			XWriter.WriteElementString("Duration", m_Duration.ToString("##0.0"))
			XWriter.WriteElementString("DurationMinutes", (60.0F * m_Duration).ToString("##0.0"))
			XWriter.WriteElementString("Progress", m_Progress.ToString("##0.00"))
			XWriter.WriteElementString("CurrentOperation", m_CurrentOperation)
			XWriter.WriteStartElement("TaskDetails")
			XWriter.WriteElementString("Status", ConvertTaskDetailStatusToString(m_TaskStatusDetail))
			XWriter.WriteElementString("Job", m_JobNumber.ToString())
			XWriter.WriteElementString("Step", m_JobStep.ToString())
			XWriter.WriteElementString("Dataset", m_Dataset)
			'TODO: Figure out how to get the most recent message
			XWriter.WriteElementString("MostRecentLogMessage", clsStatusData.MostRecentLogMessage)
			XWriter.WriteElementString("MostRecentJobInfo", m_MostRecentJobInfo)
			XWriter.WriteElementString("SpectrumCount", m_SpectrumCount.ToString())
			XWriter.WriteEndElement()	'Task details section
			XWriter.WriteEndElement()	'Task section
			XWriter.WriteEndElement()	'Root section

			'Close the document, but don't close the writer yet
			XWriter.WriteEndDocument()
			XWriter.Flush()

			'Use a streamreader to copy the XML text to a string variable
			MemStream.Seek(0, SeekOrigin.Begin)
			MemStreamReader = New System.IO.StreamReader(MemStream)
			XMLText = MemStreamReader.ReadToEnd

			MemStreamReader.Close()
			MemStream.Close()

			'Since the document is now in a string, we can close the XWriter
			XWriter.Close()
			XWriter = Nothing
			GC.Collect()
			GC.WaitForPendingFinalizers()

			'Write the output file
			Dim OutFile As StreamWriter
			Try
				OutFile = New StreamWriter(New FileStream(m_FileNamePath, FileMode.Create, FileAccess.Write, FileShare.Read))
				OutFile.WriteLine(XMLText)
				OutFile.Close()
			Catch ex As Exception
				'TODO: Figure out appropriate action
			End Try
		Catch
			'TODO: Figure out appropriate action
		End Try

		'Log to a message queue
		If m_LogToMsgQueue Then LogStatusToMessageQueue(XMLText)

	End Sub

	''' <summary>
	''' Writes the status to the message queue
	''' </summary>
	''' <param name="strStatusXML">A string contiaining the XML to write</param>
	''' <remarks></remarks>
	Protected Sub LogStatusToMessageQueue(ByVal strStatusXML As String)

		Dim Success As Boolean
		Static dtLastFailureTime As DateTime = System.DateTime.MinValue

		Try
			Dim messageSender As New MessageSender(m_MessageQueueURI, m_MessageQueueTopic, m_MgrName)

			' message queue logger sets up local message buffering (so calls to log don't block)
			' and uses message sender (as a delegate) to actually send off the messages
			Dim queueLogger As New MessageQueueLogger()
			AddHandler queueLogger.Sender, New MessageSenderDelegate(AddressOf messageSender.SendMessage)

			queueLogger.LogStatusMessage(strStatusXML)

			queueLogger.Dispose()

			messageSender.Dispose()
		Catch ex As Exception
			'TODO: Figure out how to handle error
			Success = False
		End Try
	End Sub

	''' <summary>
	''' Updates status file
	''' </summary>
	''' <param name="PercentComplete">Job completion percentage</param>
	''' <remarks>Overload to update when completion percentage is only change</remarks>
	Public Overloads Sub UpdateAndWrite(ByVal PercentComplete As Single) Implements IStatusFile.UpdateAndWrite

		m_Progress = PercentComplete
		Me.WriteStatusFile()

	End Sub

	''' <summary>
	''' Updates status file
	''' </summary>
	''' <param name="Status">Job status enum</param>
	''' <param name="PercentComplete">Job completion percentage</param>
	''' <remarks>Overload to update file when status and completion percentage change</remarks>
	Public Overloads Sub UpdateAndWrite(ByVal Status As IStatusFile.EnumTaskStatusDetail, ByVal PercentComplete As Single) Implements IStatusFile.UpdateAndWrite

		m_TaskStatusDetail = Status
		m_Progress = PercentComplete
		Me.WriteStatusFile()

	End Sub

	''' <summary>
	''' Updates status file
	''' </summary>
	''' <param name="Status">Job status enum</param>
	''' <param name="PercentComplete">Job completion percentage</param>
	''' <param name="DTACount">Number of DTA files found for Sequest analysis</param>
	''' <remarks>Overload to provide Sequest DTA count</remarks>
	Public Overloads Sub UpdateAndWrite(ByVal Status As IStatusFile.EnumTaskStatusDetail, ByVal PercentComplete As Single, _
		ByVal DTACount As Integer) Implements IStatusFile.UpdateAndWrite

		m_TaskStatusDetail = Status
		m_Progress = PercentComplete
		m_SpectrumCount = DTACount
		Me.WriteStatusFile()

	End Sub

	''' <summary>
	''' Sets status file to show mahager not running
	''' </summary>
	''' <param name="MgrError">TRUE if manager not running due to error; FALSE otherwise</param>
	''' <remarks></remarks>
	Public Sub UpdateStopped(ByVal MgrError As Boolean) Implements IStatusFile.UpdateStopped

		If MgrError Then
			m_MgrStatus = IStatusFile.EnumMgrStatus.STOPPED_ERROR
		Else
			m_MgrStatus = IStatusFile.EnumMgrStatus.STOPPED
		End If
		m_Progress = 0
		m_SpectrumCount = 0
		m_Dataset = ""
		m_JobNumber = 0
		m_Tool = ""
		m_Duration = 0
		m_TaskStatus = IStatusFile.EnumTaskStatus.NO_TASK
		m_TaskStatusDetail = IStatusFile.EnumTaskStatusDetail.NO_TASK
		Me.WriteStatusFile()

	End Sub

	''' <summary>
	''' Updates status file to show manager disabled
	''' </summary>
	''' <param name="Local">TRUE if manager disabled locally, otherwise FALSE</param>
	''' <remarks></remarks>
	Public Sub UpdateDisabled(ByVal Local As Boolean) Implements IStatusFile.UpdateDisabled

		If Local Then
			m_MgrStatus = IStatusFile.EnumMgrStatus.DISABLED_LOCAL
		Else
			m_MgrStatus = IStatusFile.EnumMgrStatus.DISABLED_MC
		End If
		m_Progress = 0
		m_SpectrumCount = 0
		m_Dataset = ""
		m_JobNumber = 0
		m_Tool = ""
		m_Duration = 0
		m_TaskStatus = IStatusFile.EnumTaskStatus.NO_TASK
		m_TaskStatusDetail = IStatusFile.EnumTaskStatusDetail.NO_TASK
		Me.WriteStatusFile()

	End Sub

	''' <summary>
	''' Initializes the status from a file, if file exists
	''' </summary>
	''' <remarks></remarks>
	Public Sub InitStatusFromFile() Implements IStatusFile.InitStatusFromFile

		Dim XmlStr As String
		Dim Doc As XmlDocument

		''Clear error queue
		'm_ErrorQueue.Clear()

		'Verify status file exists
		If Not File.Exists(m_FileNamePath) Then Exit Sub

		'Get data from status file
		Try
			XmlStr = My.Computer.FileSystem.ReadAllText(m_FileNamePath)
			'Convert to an XML document
			Doc = New XmlDocument()
			Doc.LoadXml(XmlStr)

			'Get the most recent log message
			clsStatusData.MostRecentLogMessage = Doc.SelectSingleNode("//Task/TaskDetails/MostRecentLogMessage").InnerText

			'Get the most recent job info
			m_MostRecentJobInfo = Doc.SelectSingleNode("//Task/TaskDetails/MostRecentJobInfo").InnerText

			'Get the error messsages
			For Each Xn As XmlNode In Doc.SelectNodes("//Manager/RecentErrorMessages/ErrMsg")
				clsStatusData.AddErrorMessage(Xn.InnerText)
			Next
		Catch ex As Exception
			clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.ERROR, "Exception reading status file", ex)
			Exit Sub
		End Try

	End Sub
#End Region

End Class
