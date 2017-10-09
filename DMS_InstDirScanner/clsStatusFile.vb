'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2006, Battelle Memorial Institute
' Created 06/07/2006
'
'*********************************************************************************************************

Imports System.IO
Imports System.Text
Imports System.Xml
Imports PRISM

''' <summary>
''' Provides tools for creating and updating an analysis status file
''' </summary>
Public Class clsStatusFile
    Inherits clsEventNotifier
    Implements IStatusFile

#Region "Module variables"
    'Status file name and location
    Private m_FileNamePath As String = ""

    'Manager name
    Private m_MgrName As String = ""

    'Status value
    Private m_MgrStatus As IStatusFile.EnumMgrStatus = IStatusFile.EnumMgrStatus.Stopped

    'Manager start time
    Private m_MgrStartTime As Date

    'Task start time

    'CPU utilization
    Private m_CpuUtilization As Integer = 0

    'Analysis Tool
    Private m_Tool As String = ""

    'Task status
    Private m_TaskStatus As IStatusFile.EnumTaskStatus = IStatusFile.EnumTaskStatus.No_Task

    'Task duration
    Private m_Duration As Single = 0

    'Progess (in percent)
    Private m_Progress As Single = 0

    'Current operation (freeform string)
    Private m_CurrentOperation As String = ""

    'Task status detail
    Private m_TaskStatusDetail As IStatusFile.EnumTaskStatusDetail = IStatusFile.EnumTaskStatusDetail.No_Task

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

    Private ReadOnly m_DebugLevel As Integer = 1

    'Flag to indicate if status should be logged to broker in addition to a file
    Private m_LogToMessageQueue As Boolean

    Private m_MessageSender As clsMessageSender
    Private m_QueueLogger As clsMessageQueueLogger

#End Region

#Region "Properties"
    Public Property FileNamePath As String Implements IStatusFile.FileNamePath
        Get
            Return m_FileNamePath
        End Get
        Set
            m_FileNamePath = Value
        End Set
    End Property

    Public Property MgrName As String Implements IStatusFile.MgrName
        Get
            Return m_MgrName
        End Get
        Set
            m_MgrName = Value
        End Set
    End Property

    Public Property MgrStatus As IStatusFile.EnumMgrStatus Implements IStatusFile.MgrStatus
        Get
            Return m_MgrStatus
        End Get
        Set
            m_MgrStatus = Value
        End Set
    End Property

    Public Property CpuUtilization As Integer Implements IStatusFile.CpuUtilization
        Get
            Return m_CpuUtilization
        End Get
        Set
            m_CpuUtilization = Value
        End Set
    End Property

    Public Property Tool As String Implements IStatusFile.Tool
        Get
            Return m_Tool
        End Get
        Set
            m_Tool = Value
        End Set
    End Property

    Public Property TaskStartTime As Date Implements IStatusFile.TaskStartTime

    Public Property TaskStatus As IStatusFile.EnumTaskStatus Implements IStatusFile.TaskStatus
        Get
            Return m_TaskStatus
        End Get
        Set
            m_TaskStatus = Value
        End Set
    End Property

    Public Property Duration As Single Implements IStatusFile.Duration
        Get
            Return m_Duration
        End Get
        Set
            m_Duration = Value
        End Set
    End Property

    Public Property Progress As Single Implements IStatusFile.Progress
        Get
            Return m_Progress
        End Get
        Set
            m_Progress = Value
        End Set
    End Property

    Public Property CurrentOperation As String Implements IStatusFile.CurrentOperation
        Get
            Return m_CurrentOperation
        End Get
        Set
            m_CurrentOperation = Value
        End Set
    End Property

    Public Property TaskStatusDetail As IStatusFile.EnumTaskStatusDetail Implements IStatusFile.TaskStatusDetail
        Get
            Return m_TaskStatusDetail
        End Get
        Set
            m_TaskStatusDetail = Value
        End Set
    End Property

    Public Property JobNumber As Integer Implements IStatusFile.JobNumber
        Get
            Return m_JobNumber
        End Get
        Set
            m_JobNumber = Value
        End Set
    End Property

    Public Property JobStep As Integer Implements IStatusFile.JobStep
        Get
            Return m_JobStep
        End Get
        Set
            m_JobStep = Value
        End Set
    End Property

    Public Property Dataset As String Implements IStatusFile.Dataset
        Get
            Return m_Dataset
        End Get
        Set
            m_Dataset = Value
        End Set
    End Property

    Public Property MostRecentJobInfo As String Implements IStatusFile.MostRecentJobInfo
        Get
            Return m_MostRecentJobInfo
        End Get
        Set
            m_MostRecentJobInfo = Value
        End Set
    End Property

    Public Property SpectrumCount As Integer Implements IStatusFile.SpectrumCount
        Get
            Return m_SpectrumCount
        End Get
        Set
            m_SpectrumCount = Value
        End Set
    End Property

    Public Property MessageQueueURI As String Implements IStatusFile.MessageQueueURI
        Get
            Return m_MessageQueueURI
        End Get
        Set
            m_MessageQueueURI = Value
        End Set
    End Property

    Public Property MessageQueueTopic As String Implements IStatusFile.MessageQueueTopic
        Get
            Return m_MessageQueueTopic
        End Get
        Set
            m_MessageQueueTopic = Value
        End Set
    End Property

    Public Property LogToMsgQueue As Boolean Implements IStatusFile.LogToMsgQueue
        Get
            Return m_LogToMessageQueue
        End Get
        Set
            m_LogToMessageQueue = Value
        End Set
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="FileLocation">Full path to status file</param>
    ''' <remarks></remarks>
    Public Sub New(FileLocation As String, debugLevel As Integer)
        m_FileNamePath = FileLocation
        m_MgrStartTime = DateTime.UtcNow
        m_Progress = 0
        m_SpectrumCount = 0
        m_Dataset = ""
        m_JobNumber = 0
        m_Tool = ""
        m_DebugLevel = debugLevel
    End Sub

    ''' <summary>
    ''' Converts the manager status enum to a string value
    ''' </summary>
    ''' <param name="StatusEnum">An IStatusFile.EnumMgrStatus object</param>
    ''' <returns>String representation of input object</returns>
    ''' <remarks></remarks>
    Private Function ConvertMgrStatusToString(StatusEnum As IStatusFile.EnumMgrStatus) As String

        Return StatusEnum.ToString("G")

    End Function

    ''' <summary>
    ''' Converts the task status enum to a string value
    ''' </summary>
    ''' <param name="StatusEnum">An IStatusFile.EnumTaskStatus object</param>
    ''' <returns>String representation of input object</returns>
    ''' <remarks></remarks>
    Private Function ConvertTaskStatusToString(StatusEnum As IStatusFile.EnumTaskStatus) As String

        Return StatusEnum.ToString("G")

    End Function

    ''' <summary>
    ''' Converts the manager status enum to a string value
    ''' </summary>
    ''' <param name="StatusEnum">An IStatusFile.EnumTaskStatusDetail object</param>
    ''' <returns></returns>
    ''' <remarks>String representation of input object</remarks>
    Private Function ConvertTaskDetailStatusToString(StatusEnum As IStatusFile.EnumTaskStatusDetail) As String

        Return StatusEnum.ToString("G")

    End Function


    '' <summary>
    '' Writes the status to the message queue
    '' </summary>
    '' <param name="strStatusXML">A string contiaining the XML to write</param>
    '' <remarks></remarks>
    'Protected Sub LogStatusToMessageQueueOld(ByVal strStatusXML As String)

    '	Dim Success As Boolean
    '	Static dtLastFailureTime As DateTime = DateTime.MinValue

    '	Try
    '		Dim messageSender As New MessageSender(m_MessageQueueURI, m_MessageQueueTopic, m_MgrName)

    '		' message queue logger sets up local message buffering (so calls to log don't block)
    '		' and uses message sender (as a delegate) to actually send off the messages
    '		Dim queueLogger As New MessageQueueLogger()
    '		AddHandler queueLogger.Sender, New MessageSenderDelegate(AddressOf messageSender.SendMessage)

    '		queueLogger.LogStatusMessage(strStatusXML)

    '		queueLogger.Dispose()

    '		messageSender.Dispose()
    '	Catch ex As Exception
    '		'TODO: Figure out how to handle error
    '		Success = False
    '	End Try
    'End Sub

    Protected Sub LogStatusToMessageQueue(strStatusXML As String)

        Const MINIMUM_LOG_FAILURE_INTERVAL_MINUTES As Single = 10
        Static dtLastFailureTime As DateTime = DateTime.UtcNow.Subtract(New TimeSpan(1, 0, 0))

        Try
            If m_MessageSender Is Nothing Then

                If m_DebugLevel >= 5 Then
                    LogDebug("Initializing message queue with URI '" & m_MessageQueueURI & "' and Topic '" & m_MessageQueueTopic & "'")
                End If

                m_MessageSender = New clsMessageSender(m_MessageQueueURI, m_MessageQueueTopic, m_MgrName)

                ' message queue logger sets up local message buffering (so calls to log don't block)
                ' and uses message sender (as a delegate) to actually send off the messages
                m_QueueLogger = New clsMessageQueueLogger()
                AddHandler m_QueueLogger.Sender, New MessageSenderDelegate(AddressOf m_MessageSender.SendMessage)

                If m_DebugLevel >= 3 Then
                    LogDebug("Message queue initialized with URI '" & m_MessageQueueURI & "'; posting to Topic '" & m_MessageQueueTopic & "'")
                End If

            End If

            If Not m_QueueLogger Is Nothing Then
                m_QueueLogger.LogStatusMessage(strStatusXML)
            End If

        Catch ex As Exception
            If DateTime.UtcNow.Subtract(dtLastFailureTime).TotalMinutes >= MINIMUM_LOG_FAILURE_INTERVAL_MINUTES Then
                dtLastFailureTime = DateTime.UtcNow
                OnErrorEvent("Error in clsStatusFile.LogStatusToMessageQueue (B): " & ex.Message, ex)
            End If

        End Try


    End Sub

    ''' <summary>
    ''' Writes the status file
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub WriteStatusFile() Implements IStatusFile.WriteStatusFile

        'Writes a status file for external monitor to read

        Dim XMLText As String = String.Empty

        'Set up the XML writer
        Try

            'Create a memory stream to write the document in
            Using memStream = New MemoryStream
                Using xWriter = New XmlTextWriter(memStream, Encoding.UTF8)

                    xWriter.Formatting = Formatting.Indented
                    xWriter.Indentation = 2

                    'Write the file
                    xWriter.WriteStartDocument(True)
                    'Root level element
                    xWriter.WriteStartElement("Root")
                    xWriter.WriteStartElement("Manager")
                    xWriter.WriteElementString("MgrName", m_MgrName)
                    xWriter.WriteElementString("MgrStatus", ConvertMgrStatusToString(m_MgrStatus))
                    xWriter.WriteElementString("LastUpdate", DateTime.Now().ToString)
                    xWriter.WriteElementString("LastStartTime", m_MgrStartTime.ToLocalTime().ToString())
                    xWriter.WriteElementString("CPUUtilization", m_CpuUtilization.ToString())
                    xWriter.WriteElementString("FreeMemoryMB", "0")

                    xWriter.WriteStartElement("RecentErrorMessages")
                    For Each ErrMsg As String In clsStatusData.ErrorQueue
                        xWriter.WriteElementString("ErrMsg", ErrMsg)
                    Next
                    xWriter.WriteEndElement()   'Error messages
                    xWriter.WriteEndElement()   'Manager section

                    xWriter.WriteStartElement("Task")
                    xWriter.WriteElementString("Tool", m_Tool)
                    xWriter.WriteElementString("Status", ConvertTaskStatusToString(m_TaskStatus))
                    xWriter.WriteElementString("Duration", m_Duration.ToString("##0.0"))
                    xWriter.WriteElementString("DurationMinutes", (60.0F * m_Duration).ToString("##0.0"))
                    xWriter.WriteElementString("Progress", m_Progress.ToString("##0.00"))
                    xWriter.WriteElementString("CurrentOperation", m_CurrentOperation)
                    xWriter.WriteStartElement("TaskDetails")
                    xWriter.WriteElementString("Status", ConvertTaskDetailStatusToString(m_TaskStatusDetail))
                    xWriter.WriteElementString("Job", m_JobNumber.ToString())
                    xWriter.WriteElementString("Step", m_JobStep.ToString())
                    xWriter.WriteElementString("Dataset", m_Dataset)

                    xWriter.WriteElementString("MostRecentLogMessage", clsStatusData.MostRecentLogMessage)
                    xWriter.WriteElementString("MostRecentJobInfo", m_MostRecentJobInfo)
                    xWriter.WriteElementString("SpectrumCount", m_SpectrumCount.ToString())
                    xWriter.WriteEndElement()   'Task details section
                    xWriter.WriteEndElement()   'Task section
                    xWriter.WriteEndElement()   'Root section

                    'Close the document, but don't close the writer yet
                    xWriter.WriteEndDocument()
                    xWriter.Flush()

                    'Use a streamreader to copy the XML text to a string variable
                    memStream.Seek(0, SeekOrigin.Begin)
                    Using MemStreamReader = New StreamReader(memStream)
                        XMLText = MemStreamReader.ReadToEnd

                    End Using
                End Using

            End Using

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
        If m_LogToMessageQueue Then
            ' Send the XML text to a message queue
            LogStatusToMessageQueue(XMLText)
        End If

    End Sub

    ''' <summary>
    ''' Updates status file
    ''' </summary>
    ''' <param name="PercentComplete">Job completion percentage</param>
    ''' <remarks>Overload to update when completion percentage is only change</remarks>
    Public Overloads Sub UpdateAndWrite(PercentComplete As Single) Implements IStatusFile.UpdateAndWrite

        m_Progress = PercentComplete
        Me.WriteStatusFile()

    End Sub

    ''' <summary>
    ''' Updates status file
    ''' </summary>
    ''' <param name="Status">Job status enum</param>
    ''' <param name="PercentComplete">Job completion percentage</param>
    ''' <remarks>Overload to update file when status and completion percentage change</remarks>
    Public Overloads Sub UpdateAndWrite(Status As IStatusFile.EnumTaskStatusDetail, PercentComplete As Single) Implements IStatusFile.UpdateAndWrite

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
    Public Overloads Sub UpdateAndWrite(Status As IStatusFile.EnumTaskStatusDetail, PercentComplete As Single,
        DTACount As Integer) Implements IStatusFile.UpdateAndWrite

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
    Public Sub UpdateStopped(MgrError As Boolean) Implements IStatusFile.UpdateStopped

        If MgrError Then
            m_MgrStatus = IStatusFile.EnumMgrStatus.Stopped_Error
        Else
            m_MgrStatus = IStatusFile.EnumMgrStatus.Stopped
        End If
        m_Progress = 0
        m_SpectrumCount = 0
        m_Dataset = ""
        m_JobNumber = 0
        m_Tool = ""
        m_Duration = 0
        m_TaskStatus = IStatusFile.EnumTaskStatus.No_Task
        m_TaskStatusDetail = IStatusFile.EnumTaskStatusDetail.No_Task
        Me.WriteStatusFile()

    End Sub

    ''' <summary>
    ''' Updates status file to show manager disabled
    ''' </summary>
    ''' <param name="Local">TRUE if manager disabled locally, otherwise FALSE</param>
    ''' <remarks></remarks>
    Public Sub UpdateDisabled(Local As Boolean) Implements IStatusFile.UpdateDisabled

        If Local Then
            m_MgrStatus = IStatusFile.EnumMgrStatus.Disabled_Local
        Else
            m_MgrStatus = IStatusFile.EnumMgrStatus.Disabled_MC
        End If
        m_Progress = 0
        m_SpectrumCount = 0
        m_Dataset = ""
        m_JobNumber = 0
        m_Tool = ""
        m_Duration = 0
        m_TaskStatus = IStatusFile.EnumTaskStatus.No_Task
        m_TaskStatusDetail = IStatusFile.EnumTaskStatusDetail.No_Task
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
            OnErrorEvent("Exception reading status file", ex)
            Exit Sub
        End Try

    End Sub

    Public Sub DisposeMessageQueue() Implements IStatusFile.DisposeMessageQueue
        If Not m_MessageSender Is Nothing Then
            m_QueueLogger.Dispose()
            m_MessageSender.Dispose()
        End If

    End Sub

    Public Sub LogDebug(message As String)
        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.DEBUG, message)
    End Sub

#End Region

End Class
