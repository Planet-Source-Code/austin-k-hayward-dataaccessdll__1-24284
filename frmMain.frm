VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Picture Box"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   2595
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Instructions"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add New"
      Height          =   495
      Left            =   4860
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next Record"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect using SQL"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   3480
      ScaleHeight     =   1395
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect using JET"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this project is an example of how to use the dataaccess dll
'the properties of xData listed in Command1_Click() and
'Command2_Click() must be set, and must be set in that order
'before DLL will work.

Dim xData As DataAccessDLL.clsData
Dim strAustin As String

Private Sub cmdClear_Click()

    Picture1.Cls

End Sub

Private Sub Command1_Click()

On Error GoTo Command1_Err

    Dim strSQL As String

    strSQL = "select * [PUT TABLE NAME HERE]"

    With xData

        If .xConnectionState = True Then
            .CloseConnection
        End If

        .xDatabaseName = "[PUT DATABASE PATH HERE]"
        .xCommandType = adCmdText
        .xCommandText = strSQL
        .xCursorLocation = adUseServer
        .xCursorType = adOpenStatic
        .xLockType = adLockPessimistic
        .Connect xJet

        Picture1.Print .xRecordset.Fields![FIELD NAME]

    End With

strAustin = "JET"

Exit Sub

Command1_Err:
    MsgBox "An error has occurred.  Please restart the application.", vbOKOnly, "Error"
    End

End Sub

Private Sub Command2_Click()

On Error GoTo Command2_Err

    Dim strSQL As String

    strSQL = "select * from [TABLE NAME]"

    With xData

        If .xConnectionState = True Then
            .CloseConnection
        End If

        .xDatabaseName = "[PUT DATABASE NAME HERE]"
        .xServerName = "[PUT SERVER NAME HERE]"
        .xAuthenticationMode = xSQL_Auth
        .xUserName = "sa"
        .xPassword = ""
        .xCommandType = adCmdText
        .xCommandText = strSQL
        .xCursorLocation = adUseServer
        .xCursorType = adOpenStatic
        .xLockType = adLockPessimistic
        .Connect xSQL

        Picture1.Print .xRecordset.Fields![FIELD NAME]

    End With

strAustin = "SQL"

Exit Sub

Command2_Err:
    MsgBox "An error has occurred.  Please restart the application.", vbOKOnly, "Error"
    End

End Sub

Private Sub Command3_Click()

    If strAustin = "SQL" Then
        xData.xRecordset.MoveNext
        Picture1.Print xData.xRecordset.Fields![FIELD NAME]
    ElseIf strAustin = "JET" Then
        xData.xRecordset.MoveNext
        Picture1.Print xData.xRecordset.Fields![FIELD NAME]
    End If

End Sub

Private Sub Command4_Click()

        xData.xRecordset.AddNew

End Sub

Private Sub Command5_Click()

    MsgBox xData.Instructions

End Sub

Private Sub Form_Load()

    Set xData = New DataAccessDLL.clsData

End Sub

Private Sub Form_Unload(Cancel As Integer)

    xData.Quit
    Set xData = Nothing

End Sub
