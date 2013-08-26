'File:  ShareConnector.vb
'File Contents:  ShareConnector class that connects a machine to an SMB/CIFS share
'                using a password and user name.
'Author(s):  Nathan Trimble
'Comments:  This class can be useful for connecting to SMB/CIFS shares when the use of
'SSPI isn't availabe and you don't wish to run the program as a user located on the local machine.
'It's quite comparable to the "Connect using a different user name." option in the Map Network Drive
'utility in Windows.  Much of this code came from Microsoft Knowledge Base Article - 173011.  It was
'then modified to fit our needs.
'
'Modifed: DAC (7/2/2004) -- Provided overloading for constructor, added property for share name


Public Class ShareConnector

	Private errMessage As String = ""

	Structure NETRESOURCE
		Dim dwScope As Integer
		Dim dwType As Integer
		Dim dwDisplayType As Integer
		Dim dwUsage As Integer
		Dim lpLocalName As String
		Dim lpRemoteName As String
		Dim lpComment As String
		Dim lpProvider As String
	End Structure

	Public Const NO_ERROR As Short = 0
	Public Const CONNECT_UPDATE_PROFILE As Short = &H1S
	' The following includes all the constants defined for NETRESOURCE,
	' not just the ones used in this example.
	Public Const RESOURCETYPE_DISK As Short = &H1S
	Public Const RESOURCETYPE_PRINT As Short = &H2S
	Public Const RESOURCETYPE_ANY As Short = &H0S
	Public Const RESOURCE_CONNECTED As Short = &H1S
	Public Const RESOURCE_REMEMBERED As Short = &H3S
	Public Const RESOURCE_GLOBALNET As Short = &H2S
	Public Const RESOURCEDISPLAYTYPE_DOMAIN As Short = &H1S
	Public Const RESOURCEDISPLAYTYPE_GENERIC As Short = &H0S
	Public Const RESOURCEDISPLAYTYPE_SERVER As Short = &H2S
	Public Const RESOURCEDISPLAYTYPE_SHARE As Short = &H3S
	Public Const RESOURCEUSAGE_CONNECTABLE As Short = &H1S
	Public Const RESOURCEUSAGE_CONTAINER As Short = &H2S
	' Error Constants:
	Public Const ERROR_ACCESS_DENIED As Short = 5
	Public Const ERROR_ALREADY_ASSIGNED As Short = 85
	Public Const ERROR_BAD_DEV_TYPE As Short = 66
	Public Const ERROR_BAD_DEVICE As Short = 1200
	Public Const ERROR_BAD_NET_NAME As Short = 67
	Public Const ERROR_BAD_PROFILE As Short = 1206
	Public Const ERROR_BAD_PROVIDER As Short = 1204
	Public Const ERROR_BUSY As Short = 170
	Public Const ERROR_CANCELLED As Short = 1223
	Public Const ERROR_CANNOT_OPEN_PROFILE As Short = 1205
	Public Const ERROR_DEVICE_ALREADY_REMEMBERED As Short = 1202
	Public Const ERROR_EXTENDED_ERROR As Short = 1208
	Public Const ERROR_INVALID_PASSWORD As Short = 86
	Public Const ERROR_NO_NET_OR_BAD_PATH As Short = 1203

	Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

	Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Integer, ByVal fForce As Integer) As Integer

	Private mNetResource As NETRESOURCE
	Private mUsername As String
	Private mPassword As String
	Private mShareName As String = ""

	Public Sub New(ByVal share As String, ByVal userName As String, ByVal userPwd As String)

		'Overloaded version to allow providing share name in constructor
		DefineShareName(share)
		RealNew(userName, userPwd)

	End Sub

	Public Sub New(ByVal userName As String, ByVal userPwd As String)

		'Overloaded version requiring sharename to be specified as property
		RealNew(userName, userPwd)

	End Sub

	Private Sub RealNew(ByVal userName As String, ByVal userPwd As String)

		'Actual constructor called by assorted overloaded versions
		mUsername = userName
		mPassword = userPwd
		mNetResource.lpRemoteName = mShareName
		mNetResource.dwType = RESOURCETYPE_DISK
		mNetResource.dwScope = RESOURCE_GLOBALNET
		mNetResource.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
		mNetResource.dwUsage = RESOURCEUSAGE_CONNECTABLE

	End Sub

	Public Property Share() As String
		Get
			Return mShareName
		End Get
		Set(ByVal Value As String)
			DefineShareName(Value)
			mNetResource.lpRemoteName = mShareName
		End Set
	End Property

	Public Function Connect(ByVal Share As String) As Boolean

		'Connects to specified share using specified account/password
		'Overload to allow specification of share name in function call
		DefineShareName(Share)
		mNetResource.lpRemoteName = mShareName
		Return RealConnect()

	End Function

	Public Function Connect() As Boolean

		'Connects to specified share using specified account/password
		'Overload requiring specification of share name prior to function call
		If mNetResource.lpRemoteName = "" Then
			ErrorMessage = "Share name not specified"
			Return False
		End If
		Return RealConnect()

	End Function

	Private Sub DefineShareName(ByVal share As String)
		If share.EndsWith("\") Then
			mShareName = share.TrimEnd("\"c)
		Else
			mShareName = share
		End If
	End Sub

	Private Function RealConnect() As Boolean

		'Connects to specified share using specified account/password
		'This is the function that actually does the connection based on the setup from the overloaded functions
		Dim errorNum As Integer

		errorNum = WNetAddConnection2(mNetResource, mPassword, mUsername, 0)
		If errorNum = NO_ERROR Then
			Debug.WriteLine("Connected.")
			Return True
		Else
			ErrorMessage = errorNum.ToString()
			Debug.WriteLine("Got error: " & errorNum)
			Return False
		End If

	End Function

	Public Function Disconnect() As Boolean
		Dim errorNum As Integer = WNetCancelConnection2(Me.mNetResource.lpRemoteName, 0, CInt(True))
		If errorNum = NO_ERROR Then
			Debug.WriteLine("Disconnected.")
			Return True
		Else
			ErrorMessage = errorNum.ToString()
			Debug.WriteLine("Got error: " & errorNum)
			Return False
		End If
	End Function

	Public Property ErrorMessage() As String
		Get
			Return errMessage
		End Get
		Set(ByVal value As String)
			errMessage = value
		End Set
	End Property

End Class