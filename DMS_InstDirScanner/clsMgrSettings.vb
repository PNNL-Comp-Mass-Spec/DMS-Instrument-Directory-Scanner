'*********************************************************************************************************
' Written by Dave Clark for the US Department of Energy
' Pacific Northwest National Laboratory, Richland, WA
' Copyright 2007, Battelle Memorial Institute
' Created 07/27/2009
'
'*********************************************************************************************************
Imports System.Xml

#Region "Interfaces"
Public Interface IMgrParams
    Function GetParam(ItemKey As String) As String
    Function GetParam(ItemKey As String, valueIfMissing As String) As String
    Function GetParam(ItemKey As String, valueIfMissing As Integer) As Integer
    Sub SetParam(ItemKey As String, ItemValue As String)
End Interface
#End Region

''' <summary>
''' Class for loading, storing and accessing manager parameters.
''' </summary>
''' <remarks>
''' Loads initial settings from local config file, then checks to see if remainder of settings should be
''' loaded or manager set to inactive. If manager active, retrieves remainder of settings from manager
''' parameters database.
''' </remarks>
Public Class clsMgrSettings
    Implements IMgrParams

#Region "Module variables"
    Private m_MgrParams As Dictionary(Of String, String)
    Private m_ErrMsg As String = ""
#End Region

#Region "Properties"
    Public ReadOnly Property ErrMsg As String
        Get
            Return m_ErrMsg
        End Get
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        If Not LoadSettings() Then
            Throw New ApplicationException("Unable to initialize manager settings class")
        End If

    End Sub

    ''' <summary>
    ''' Loads manager settings from config file and database
    ''' </summary>
    ''' <returns>True if successful; False on error</returns>
    ''' <remarks></remarks>
    Public Function LoadSettings() As Boolean

        m_ErrMsg = ""

        'Get settings from config file
        m_MgrParams = LoadMgrSettingsFromFile()

        'Test the settings retrieved from the config file
        If Not CheckInitialSettings(m_MgrParams) Then
            'Error logging was already handled by CheckInitialSettings
            Return False
        End If

        'Determine if manager is deactivated locally
        If Not CBool(m_MgrParams("MgrActive_Local")) Then
            LogWarning("Manager deactivated locally")
            m_ErrMsg = "Manager deactivated locally"
            Return False
        End If

        'Get remaining settings from database
        If Not LoadMgrSettingsFromDB(m_MgrParams) Then
            'Error logging handled by LoadMgrSettingsFromDB
            Return False
        End If

        'No problems found
        Return True

    End Function

    ''' <summary>
    ''' Loads the initial settings from application config file
    ''' </summary>
    ''' <returns>String dictionary containing initial settings if suceessful; NOTHING on error</returns>
    ''' <remarks></remarks>
    Private Function LoadMgrSettingsFromFile() As Dictionary(Of String, String)

        'Load initial settings into string dictionary for return
        Dim mgrParams As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        My.Settings.Reload()
        'Manager config db connection string
        mgrParams.Add("MgrCnfgDbConnectStr", My.Settings.MgrCnfgDbConnectStr)

        'Manager active flag
        mgrParams.Add("MgrActive_Local", My.Settings.MgrActive_Local.ToString())

        'Manager name
        mgrParams.Add("MgrName", My.Settings.MgrName)

        'Default settings in use flag
        mgrParams.Add("UsingDefaults", My.Settings.UsingDefaults.ToString())

        Return mgrParams

    End Function

    ''' <summary>
    ''' Tests initial settings retrieved from config file
    ''' </summary>
    ''' <param name="mgrParams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckInitialSettings(mgrParams As IReadOnlyDictionary(Of String, String)) As Boolean

        'Verify manager settings dictionary exists
        If mgrParams Is Nothing Then
            LogError("clsMgrSettings.CheckInitialSettings(); Manager parameters dictionary is null")
            Return False
        End If

        'Verify intact config file was found
        Dim strUsingDefaults = ""
        If Not mgrParams.TryGetValue("UsingDefaults", strUsingDefaults) Then
            LogError("clsMgrSettings.CheckInitialSettings(); Manager parameter 'UsingDefaults' is not defined")
            Return False
        End If

        Dim usingDefaults As Boolean
        If Not Boolean.TryParse(strUsingDefaults, usingDefaults) Then
            LogError("clsMgrSettings.CheckInitialSettings(); Manager parameter 'UsingDefaults' must be True or False, not " & strUsingDefaults)
            Return False
        End If

        If usingDefaults Then
            LogError("clsMgrSettings.CheckInitialSettings(); Config file problem: default settings being used (UsingDefaults is true)")
            Return False
        End If

        'No problems found
        Return True

    End Function

    ''' <summary>
    ''' Gets remaining manager config settings from config database;
    ''' Overload to use module-level string dictionary when calling from external method
    ''' </summary>
    ''' <returns>True for success; False for error</returns>
    ''' <remarks></remarks>
    Public Overloads Function LoadMgrSettingsFromDB() As Boolean

        Return LoadMgrSettingsFromDB(m_MgrParams)

    End Function


    ''' <summary>
    ''' Gets remaining manager config settings from config database
    ''' </summary>
    ''' <param name="mgrSettings">String dictionary containing parameters that have been loaded so far</param>
    ''' <returns>True for success; False for error</returns>
    ''' <remarks></remarks>
    Public Overloads Function LoadMgrSettingsFromDB(mgrSettings As Dictionary(Of String, String)) As Boolean

        'Requests job parameters from database. Input string specifies view to use. Performs retries if necessary.

        Dim mgrName = m_MgrParams("MgrName")

        Dim columns = New List(Of String) From {
                "ParameterName",
                "ParameterValue"
                }

        Dim sqlStr As String =
                " SELECT " & String.Join(",", columns) &
                " FROM V_MgrParams " &
                " WHERE ManagerName = '" & mgrName & "'"

        Dim connectionString = mgrSettings("MgrCnfgDbConnectStr")
        If String.IsNullOrWhiteSpace(connectionString) Then
            LogError("Connection string is empty; cannot retrieve manager parameters")
            Return False
        End If

        Dim dbTools = New PRISM.clsDBTools(connectionString)
        AddHandler dbTools.ErrorEvent, AddressOf DBToolsErrorHandler
        AddHandler dbTools.WarningEvent, AddressOf DBToolsWarningHandler

        ' Get a table holding the parameters for one manager

        Dim lstResults As List(Of List(Of String)) = Nothing
        Dim success = dbTools.GetQueryResults(sqlStr, lstResults, "LoadMgrSettingsFromDB")

        If Not success Then
            LogError("Unable to retrieve manager settings from the database for manager " & mgrName)
            Return False
        End If

        'Verify at least one row returned
        If lstResults.Count < 1 Then
            LogError("clsMgrSettings.LoadMgrSettingsFromDB; manager settings not found in V_MgrParams for manager " & mgrName)
            Return False
        End If

        Dim colMapping = dbTools.GetColumnMapping(columns)

        'Fill a string dictionary with the manager parameters that have been found
        Try
            For Each result In lstResults

                Dim paramKey = dbTools.GetColumnValue(result, colMapping, "ParameterName")
                Dim paramVal = dbTools.GetColumnValue(result, colMapping, "ParameterValue")

                If m_MgrParams.ContainsKey(paramKey) Then
                    m_MgrParams(paramKey) = paramVal
                Else
                    m_MgrParams.Add(paramKey, paramVal)
                End If
            Next

            Return True
        Catch ex As Exception
            LogError("clsMgrSettings.LoadMgrSettingsFromDB; Exception filling string dictionary from table: " & ex.Message)
            Return False
        End Try

    End Function

    Private Sub LogError(message As String)
        PRISM.ConsoleMsgUtils.ShowError(message)
        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogSystem, clsLogTools.LogLevels.ERROR, message)
    End Sub

    Private Sub LogWarning(message As String)
        PRISM.ConsoleMsgUtils.ShowWarning(message)
        clsLogTools.WriteLog(clsLogTools.LoggerTypes.LogFile, clsLogTools.LogLevels.WARN, message)
    End Sub

    ''' <summary>
    ''' Gets a parameter from the parameters string dictionary
    ''' </summary>
    ''' <param name="ItemKey">Key name for item</param>
    ''' <returns>String value associated with specified key</returns>
    ''' <remarks>Returns Nothing if key isn't found</remarks>
    Public Function GetParam(ItemKey As String, valueIfMissing As String) As String Implements IMgrParams.GetParam

        If Not m_MgrParams.ContainsKey(ItemKey) Then
            Return valueIfMissing
        End If

        Return m_MgrParams.Item(ItemKey)

    End Function

    ''' <summary>
    ''' Gets a parameter from the parameters string dictionary
    ''' </summary>
    ''' <param name="ItemKey">Key name for item</param>
    ''' <returns>String value associated with specified key</returns>
    ''' <remarks>Returns Nothing if key isn't found</remarks>
    Public Function GetParam(ItemKey As String, valueIfMissing As Integer) As Integer Implements IMgrParams.GetParam

        If Not m_MgrParams.ContainsKey(ItemKey) Then
            Return valueIfMissing
        End If

        Dim strValue = m_MgrParams.Item(ItemKey)
        Dim value As Integer
        If Integer.TryParse(strValue, value) Then
            Return value
        End If

        Return valueIfMissing

    End Function

    ''' <summary>
    ''' Gets a parameter from the parameters string dictionary
    ''' </summary>
    ''' <param name="ItemKey">Key name for item</param>
    ''' <returns>String value associated with specified key</returns>
    ''' <remarks>Returns Nothing if key isn't found</remarks>
    Public Function GetParam(ItemKey As String) As String Implements IMgrParams.GetParam

        Return m_MgrParams.Item(ItemKey)

    End Function

    ''' <summary>
    ''' Sets a parameter in the parameters string dictionary
    ''' </summary>
    ''' <param name="ItemKey">Key name for the item</param>
    ''' <param name="ItemValue">Value to assign to the key</param>
    ''' <remarks></remarks>
    Public Sub SetParam(ItemKey As String, ItemValue As String) Implements IMgrParams.SetParam

        m_MgrParams.Item(ItemKey) = ItemValue

    End Sub

    ''' <summary>
    ''' Gets a collection representing all keys in the parameters string dictionary
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAllKeys() As ICollection

        Return m_MgrParams.Keys

    End Function

    ''' <summary>
    ''' Writes specfied value to an application config file.
    ''' </summary>
    ''' <param name="Key">Name for parameter (case sensitive)</param>
    ''' <param name="Value">New value for parameter</param>
    ''' <returns>TRUE for success; FALSE for error (ErrMsg property contains reason)</returns>
    ''' <remarks>This bit of lunacy is needed because MS doesn't supply a means to write to an app config file</remarks>
    Public Function WriteConfigSetting(Key As String, Value As String) As Boolean

        m_ErrMsg = ""

        'Load the config document
        Dim MyDoc As XmlDocument = LoadConfigDocument()
        If MyDoc Is Nothing Then
            'Error message has already been produced by LoadConfigDocument
            Return False
        End If

        'Retrieve the settings node
        Dim MyNode As XmlNode = MyDoc.SelectSingleNode("//applicationSettings")

        If MyNode Is Nothing Then
            m_ErrMsg = "clsMgrSettings.WriteConfigSettings; appSettings node not found"
            Return False
        End If

        Try
            'Select the eleement containing the value for the specified key containing the key
            Dim MyElement = CType(MyNode.SelectSingleNode(String.Format("//setting[@name='{0}']/value", Key)), XmlElement)
            If MyElement IsNot Nothing Then
                'Set key to specified value
                MyElement.InnerText = Value
            Else
                'Key was not found
                m_ErrMsg = "clsMgrSettings.WriteConfigSettings; specified key not found: " & Key
                Return False
            End If
            MyDoc.Save(GetConfigFilePath())
            Return True
        Catch ex As Exception
            m_ErrMsg = "clsMgrSettings.WriteConfigSettings; Exception updating settings file: " & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Loads an app config file for changing parameters
    ''' </summary>
    ''' <returns>App config file as an XML document if successful; NOTHING on failure</returns>
    ''' <remarks></remarks>
    Private Function LoadConfigDocument() As XmlDocument

        Try
            Dim myDoc = New XmlDocument
            myDoc.Load(GetConfigFilePath)
            Return myDoc
        Catch ex As Exception
            m_ErrMsg = "clsMgrSettings.LoadConfigDocument; Exception loading settings file: " & ex.Message
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Specifies the full name and path for the application config file
    ''' </summary>
    ''' <returns>String containing full name and path</returns>
    ''' <remarks></remarks>
    Private Function GetConfigFilePath() As String

        Return clsMainProcess.GetAppPath() & ".config"

    End Function

    ''' <summary>
    ''' Converts a database output object that could be dbNull to a string
    ''' </summary>
    ''' <param name="InpObj"></param>
    ''' <returns>String equivalent of object; empty string if object is dbNull</returns>
    ''' <remarks></remarks>
    Protected Function DbCStr(InpObj As Object) As String

        'If input object is DbNull, returns "", otherwise returns String representation of object
        If InpObj Is DBNull.Value Then
            Return ""
        Else
            Return CStr(InpObj)
        End If

    End Function

    Private Sub DBToolsErrorHandler(message As String, ex As Exception)
        LogError(message)
    End Sub

    Private Sub DBToolsWarningHandler(message As String)
        LogWarning(message)
    End Sub

#End Region

End Class