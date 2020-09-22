VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************
'*                             *
'*     Database access DLL     *
'*       Copyright ©2001       *
'*      Austin K. Hayward      *
'*                             *
'*******************************

Option Explicit

Dim xConnection As ADODB.Connection
Public xRecordset As ADODB.Recordset

Public Enum xAccessType
    xJet = 0
    xSQL = 1
End Enum

Public Enum xAuthenMode
    xNT_Auth = 0
    xSQL_Auth
End Enum

Private mvarxConnectionState As Boolean
Private mvarxServerName As String
Private mvarxDatabaseName As String
Private mvarxUserName As String
Private mvarxPassword As String
Private mvarxAuthenticationMode As xAuthenMode
Private mvarxCommandType As ADODB.CommandTypeEnum
Private mvarxCommandText As String
Private mvarxCursorLocation As ADODB.CursorLocationEnum
Private mvarxCursorType As ADODB.CursorTypeEnum
Private mvarxLockType As ADODB.LockTypeEnum

Public Property Let xAuthenticationMode(ByVal vData As xAuthenMode)
    mvarxAuthenticationMode = vData
End Property

Public Property Get xAuthenticationMode() As xAuthenMode
    xAuthenticationMode = mvarxAuthenticationMode
End Property

Public Property Let xPassword(ByVal vData As String)
    mvarxPassword = vData
End Property

Public Property Get xPassword() As String
    xPassword = mvarxPassword
End Property

Public Property Let xUserName(ByVal vData As String)
    mvarxUserName = vData
End Property

Public Property Get xUserName() As String
    xUserName = mvarxUserName
End Property

Public Property Let xLockType(ByVal vData As ADODB.LockTypeEnum)
    mvarxLockType = vData
End Property

Public Property Get xLockType() As ADODB.LockTypeEnum
    xLockType = mvarxLockType
End Property

Public Property Let xCursorType(ByVal vData As ADODB.CursorTypeEnum)
    mvarxCursorType = vData
End Property

Public Property Get xCursorType() As ADODB.CursorTypeEnum
    xCursorType = mvarxCursorType
End Property

Public Property Let xCursorLocation(ByVal vData As ADODB.CursorLocationEnum)
    mvarxCursorLocation = vData
End Property

Public Property Get xCursorLocation() As ADODB.CursorLocationEnum
    xCursorLocation = mvarxCursorLocation
End Property

Public Property Let xCommandText(ByVal vData As String)
    mvarxCommandText = vData
End Property

Public Property Get xCommandText() As String
    xCommandText = mvarxCommandText
End Property

Public Property Let xConnectionState(ByVal vData As Boolean)
    mvarxConnectionState = vData
End Property

Public Property Get xConnectionState() As Boolean
    xConnectionState = mvarxConnectionState
End Property

Public Property Let xDatabaseName(ByVal vData As String)
    mvarxDatabaseName = vData
End Property

Public Property Get xDatabaseName() As String
    xDatabaseName = mvarxDatabaseName
End Property

Public Property Let xServerName(ByVal vData As String)
    mvarxServerName = vData
End Property

Public Property Get xServerName() As String
    xServerName = mvarxServerName
End Property

Public Property Let xCommandType(ByVal vData As ADODB.CommandTypeEnum)
    mvarxCommandType = vData
End Property

Public Property Get xCommandType() As ADODB.CommandTypeEnum
    xCommandType = mvarxCommandType
End Property

Public Sub Connect(aType As xAccessType)

On Error GoTo Connect_Err
    
    Select Case aType
        Case Is = 0
            Call ConnectJet(xDatabaseName, xCommandType, xCommandText, xCursorLocation, xCursorType, xLockType)
        Case Is = 1
            Call ConnectSQL(xServerName, xDatabaseName, xAuthenticationMode, xUserName, xPassword, xCommandType, xCommandText, xCursorLocation, xCursorType, xLockType)
        Case Else
            MsgBox "An error was encountered.  Please restart the application.", vbOKOnly, "Error"
    End Select

Exit Sub

Connect_Err:

    MsgBox "An error was encountered.  Please restart the application.", vbOKOnly, "Error"

End Sub

Private Sub ConnectJet(aDatabaseName As String, aCommandType As ADODB.CommandTypeEnum, aCommandText As String, aCursorLocation As ADODB.CursorLocationEnum, aCursorType As ADODB.CursorTypeEnum, aLockType As ADODB.LockTypeEnum)

On Error GoTo ConnectJet_Err

    Set xConnection = New ADODB.Connection
    Set xRecordset = New ADODB.Recordset

    xDatabaseName = aDatabaseName

    Dim xConnectionString As String

    xConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & aDatabaseName & ";Persist Security Info=False"

    With xConnection
        .Open xConnectionString
    End With

    If xRecordset.State = adStateOpen Then
        xRecordset.Close
    End If

    With xRecordset
        .CursorLocation = aCursorLocation
        .CursorType = aCursorType
        .LockType = aLockType
        .Open xCommandText, xConnection, , , xCommandType
    End With

    xConnectionState = True

Exit Sub

ConnectJet_Err:
    MsgBox Err.Source & " - " & Err.Description
    Resume Next

End Sub

Private Sub ConnectSQL(aServerName As String, aDatabaseName As String, aAuthenMode As xAuthenMode, aUserName As String, aPassword As String, aCommandType As ADODB.CommandTypeEnum, aCommandText As String, aCursorLocation As ADODB.CursorLocationEnum, aCursorType As ADODB.CursorTypeEnum, aLockType As ADODB.LockTypeEnum)

On Error GoTo ConnectSQL_Err

    Set xConnection = New ADODB.Connection
    Set xRecordset = New ADODB.Recordset

    aDatabaseName = xDatabaseName
    aServerName = xServerName
    aUserName = xUserName
    aPassword = xPassword

    Dim xConnectionString As String

    If xAuthenticationMode = 0 Then
        xConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & aDatabaseName & ";Data Source=" & aServerName
    ElseIf xAuthenticationMode = 1 Then
        xConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & aUserName & ";Password=" & aPassword & ";Initial Catalog=" & aDatabaseName & ";Data Source=" & aServerName
    Else
        MsgBox "Please choose an authentication mode.", vbOKOnly, "DataAccessDLL"
        Exit Sub
    End If

    With xConnection
        .ConnectionTimeout = 30
        .Open xConnectionString
    End With

    If xRecordset.State = adStateOpen Then
        xRecordset.Close
    End If

    With xRecordset
        .CursorLocation = aCursorLocation
        .CursorType = aCursorType
        .LockType = aLockType
        .Open xCommandText, xConnection, , , xCommandType
    End With
    
    xConnectionState = True

Exit Sub

ConnectSQL_Err:
    MsgBox Err.Source & " - " & Err.Description
    Resume Next

End Sub

Private Sub Class_Initialize()

On Error GoTo Class_Initialize_Err

    xConnectionState = False

Exit Sub

Class_Initialize_Err:

    MsgBox "An error occurred.  Check with your system administrator.", vbOKOnly, "Error"

End Sub

Public Sub Quit()

    Set xConnection = Nothing
    Set xRecordset = Nothing
    xConnectionState = False
    xServerName = ""
    xDatabaseName = ""
    xUserName = ""
    xPassword = ""
    xCommandText = ""

End Sub

Public Sub CloseConnection()

    On Error Resume Next

    If xConnection.State = adStateOpen Then
        xConnection.Close
        xConnectionState = False
    End If

End Sub

Public Function Instructions() As String

    Dim xInstructions As String

    xInstructions = "JET -" & vbLf & vbLf & _
                    "xDatabaseName" & vbLf & _
                    "xCommandType" & vbLf & _
                    "xCommandText" & vbLf & _
                    "xCursorLocation" & vbLf & _
                    "xCursorType" & vbLf & _
                    "xLockType" & vbLf & _
                    "Connect" & vbLf & vbLf & vbLf & _
                    "SQL -" & vbLf & vbLf & _
                    "xDatabaseName" & vbLf & _
                    "xServerName" & vbLf & _
                    "xAuthenticationMode" & vbLf & _
                    "xUserName" & vbLf & _
                    "xPassword" & vbLf & _
                    "xCommandType" & vbLf & _
                    "xCommandText" & vbLf & _
                    "xCursorLocation" & vbLf & _
                    "xCursorType" & vbLf & _
                    "xLockType" & vbLf & _
                    "Connect" & vbLf & vbLf & vbLf & _
                    "Must be in the above order -" & vbLf & _
                    "Austin"

    Instructions = xInstructions

End Function
