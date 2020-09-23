VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "LogOn"
   ClientHeight    =   4665
   ClientLeft      =   4260
   ClientTop       =   3660
   ClientWidth     =   7380
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin EasyStat.UPB UPB1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      ForeColor       =   33023
      FollowColor     =   16777215
      Borders         =   6
      ForeSpeed       =   8
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   480
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "1721975"
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0ECA
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   3240
      Picture         =   "Form2.frx":0ED0
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000000&
      Caption         =   "Installation Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Developed by : Tatai Basu "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   6615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agricultural Statistics  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EASY AGRICULTURAL STATISTICS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   5520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2760
      Width           =   5175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim isp(500) As String
Dim ins As Integer
Dim isp1(500) As String
Dim ins1 As Integer

Open "C:\WINDOWS\Hi.txt" For Input As #9
ins = 0
Do
ins = ins + 1
On Error Resume Next
Input #9, isp(ins)


Text1.Text = isp(ins)
Loop Until EOF(9)
Close #9

Open "C:\WINDOWS\HiDate.txt" For Input As #10
ins1 = 0
Do
ins1 = ins1 + 1
On Error Resume Next
Input #10, isp1(ins1)


Text2.Text = isp1(ins1)
Loop Until EOF(10)
Close #10
End Sub

Private Sub Timer1_Timer()
Dim d As Double
Static sec
sec = sec + 1

If sec = 6 Then
UPB1.Enabled = True
Label1.Caption = "Veryfying Piracy....."
End If

If sec >= 40 Then
UPB1.ForeColor = vbvlue
UPB1.FollowColor = vbYellow
d = StrComp(Text1.Text, Text3.Text, vbTextCompare)
Label1.Caption = "Veryfying Computer...."

End If
If sec >= 60 Then
UPB1.ForeColor = vbGreen
UPB1.FollowColor = vbBlue
If d = 0 Then
Label1.Caption = "Original Copy -- Loading...."
MDIForm1.Show
MDIForm1.Frame1.Visible = False
MDIForm1.Text1.SetFocus
End If
ElseIf d = -1 Then
Label1.Caption = "Pyrated Copy -- Contact Amal Chakraborty (9434339500)"
MsgBox "Pyrated Copy -- Contact tatai.basu (9434339500)"
Unload Me
End If
If sec = 65 Then
Unload Me
End If

End Sub
