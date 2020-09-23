VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "This is the client!"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "frmClient.frx":0000
      Left            =   0
      List            =   "frmClient.frx":0002
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "frmClient.frx":0004
      Left            =   0
      List            =   "frmClient.frx":0006
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock wskClient 
      Left            =   120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FirstLoad()
'Get the title
Dim sTitle As String

sTitle = Me.Caption & "  " & wskClient.LocalHostName & _
            " - " & wskClient.LocalIP

Me.Caption = sTitle
End Sub

Private Sub Command1_Click()
Command1.Enabled = Not (Command1.Enabled)
'Make a connection
With wskClient
    .RemoteHost = "127.0.0.1"
    .RemotePort = 1002
    .Bind 1001
End With

End Sub

Private Sub Command2_Click()
List2.Clear
End Sub

Private Sub Form_Load()
'Add commands to the list1
FirstLoad
List1.AddItem "GET_REPORT"
List1.AddItem "CLEAR_REPORT"
End Sub

Private Sub List1_DblClick()
'Send the text of the command selected in list1
wskClient.SendData List1.List(List1.ListIndex)
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
'Get a response and fill the list2
Dim sData As String

wskClient.GetData sData
List2.AddItem sData
End Sub

