VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MainForm"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   15240
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   10095
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Evaluation Wing "
            TextSave        =   "Evaluation Wing "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "4:04 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "10/15/2009"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   13124
            MinWidth        =   13124
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   3422
            MinWidth        =   3422
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10095
      Left            =   0
      ScaleHeight     =   10035
      ScaleWidth      =   15180
      TabIndex        =   1
      Top             =   0
      Width           =   15240
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "LOG ON PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9240
         TabIndex        =   4
         Top             =   480
         Width           =   5535
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "EXIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4200
            Picture         =   "MDIForm1.frx":0ECA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   2760
            Picture         =   "MDIForm1.frx":1794
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   240
            PasswordChar    =   "a"
            TabIndex        =   5
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Password : 13102k"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MENU"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   9735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton Command5 
            BackColor       =   &H0000FFFF&
            Caption         =   "EXIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            Picture         =   "MDIForm1.frx":1A9E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   8760
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0FF&
            Caption         =   "POTATO YEILDS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            Picture         =   "MDIForm1.frx":2368
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agricultural Statistics  "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   5520
         TabIndex        =   3
         Top             =   3720
         Width           =   5295
      End
      Begin VB.Image Image2 
         Height          =   1365
         Left            =   5880
         Picture         =   "MDIForm1.frx":27AA
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   10140
         Left            =   0
         Picture         =   "MDIForm1.frx":6017
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   15180
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Index           =   11
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
         Index           =   12
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
If Text1.Text = "13102k" Then
Frame1.Visible = True
Frame2.Visible = False
Else
MsgBox "Wrong Password"
Text1.Text = ""
Text1.SetFocus

Exit Sub
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command3.SetFocus
End If
End Sub
