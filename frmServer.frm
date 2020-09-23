VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H00C0C0C0&
   Caption         =   "This is the UDP server!"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Clear"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Ok Buy it"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   11
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "frmServer.frx":0000
      Left            =   2280
      List            =   "frmServer.frx":0002
      TabIndex        =   9
      Top             =   2280
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "frmServer.frx":0004
      Left            =   240
      List            =   "frmServer.frx":0006
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   3240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   600
   End
   Begin VB.Label Label6 
      Caption         =   "To Buy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   10
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label5 
      Caption         =   "Items to Sell"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   " 00.00 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "Price : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "The Big Hamster - Pets Shop ;)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1005
      TabIndex        =   4
      Top             =   120
      Width           =   4365
   End
   Begin VB.Label lblDateTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HH:MM:SS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Time : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblDateTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DD-MM-YYYY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   690
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is an array to use with 3 kinds of pets
'and a integer variant to keep the count of
'the selling

Dim PetAdmin(1 To 3) As PET_TYPE
Dim iTotal As Integer

Private Sub SetDateTime(Index As Integer)
'This sub get the index to know which
'command and label will be used

Select Case Index
'lblDateTime(index) is label a matrix
'from 0 to 1
    Case 0
        lblDateTime(Index).Caption = Format(Date, "DD-MM-YYYY")
    
    Case 1
        lblDateTime(Index).Caption = Format(Time, "HH:MM:SS")
    
End Select
End Sub

Private Sub cmdAction_Click(Index As Integer)
'Get the total or clear the list2

Select Case Index
    Case 0
        List2.AddItem "Total to pay is: " & iTotal
        List2.AddItem " "
        iTotal = 0
        
    Case 1
        iTotal = 0
        List2.Clear
        
End Select
End Sub

Private Sub Form_Load()
Call FirstLoad

'Get the connection UDP
'Remember, when you are using UDP
'connections you have to cchange the
'winsocks protocol property to scKUDPProtocol

With wskServer
'Just give the remote IP
    .RemoteHost = "127.0.0.1"
    
'The remote port to connect to..
    .RemotePort = 1001
    
'And the port where you are getting data pks
    .Bind 1002
    
End With

'Show Client
frmClient.Show
End Sub

Private Sub List1_Click()
'Get the price from the type using list1 as index
lblPrice = PetAdmin(List1.ListIndex + 1).Price
End Sub

Private Sub List1_DblClick()
'Add items to list2 keeping the Sells total amount
Dim Sdate, sTime As String
Sdate = lblDateTime(0).Caption
sTime = lblDateTime(1).Caption
iTotal = iTotal + PetAdmin(List1.ListIndex + 1).Price
List2.AddItem PetAdmin(List1.ListIndex + 1).Name & _
                " Cost " & PetAdmin(List1.ListIndex + 1).Price & _
                " - " & Sdate & " - " & sTime
End Sub

Private Sub Timer1_Timer()
'You know the time
SetDateTime (1)
End Sub

Private Sub FirstLoad()
'Loads the first caption in labels
Dim x As Integer, sTitle As String

sTitle = Me.Caption & "  " & wskServer.LocalHostName & _
            " - " & wskServer.LocalIP
'Local IP = as show in title bar or 127.0.0.1

Me.Caption = sTitle

For x = 0 To 1
'SetDateTime(Index on label from 0 to 1)
    SetDateTime (x)
Next

PetAdmin(1).Name = "Hamster"
PetAdmin(1).Price = 2
PetAdmin(2).Name = "Cat"
PetAdmin(2).Price = 25.12
PetAdmin(3).Name = "Dog"
PetAdmin(3).Price = 52.73

For x = 1 To 3
List1.AddItem PetAdmin(x).Name
Next

End Sub

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)
'This is the event wich can give you all
'the information sended by the other app

'Just use a variable, in this case a string to
'get text

Dim sDatos As String

'Use the variable to get the info everytime
'your app get a data
wskServer.GetData sDatos

Select Case sDatos
    Case "GET_REPORT"
        'Fill the list2
        Dim iItems As Integer
        For iItems = 0 To List2.ListCount - 1
        wskServer.SendData List2.List(iItems)
        Next
    
    Case "CLEAR_REPORT"
        'Clear the list2
        List2.Clear

End Select
End Sub

