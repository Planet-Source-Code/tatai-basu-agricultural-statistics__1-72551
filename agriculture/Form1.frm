VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "PotatoSheet"
   ClientHeight    =   11400
   ClientLeft      =   -75
   ClientTop       =   -135
   ClientWidth     =   15360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11400
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command10 
      BackColor       =   &H0000FF00&
      Caption         =   "NEW G.P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12960
      Picture         =   "Form1.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "DATA INPUT TABLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4575
      Left            =   0
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   12975
      Begin VB.Frame Frame14 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Caption         =   "Frame14"
         Height          =   1095
         Left            =   5400
         TabIndex        =   136
         Top             =   2880
         Width           =   7455
         Begin VB.CommandButton Command8 
            BackColor       =   &H0080FF80&
            Caption         =   "Print Report"
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
            Left            =   4560
            Picture         =   "Form1.frx":130C
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H0000FF00&
            Caption         =   "Calculator"
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
            Left            =   3120
            Picture         =   "Form1.frx":21D6
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H0080FFFF&
            Caption         =   "Total"
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
            Left            =   1680
            Picture         =   "Form1.frx":2940
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H0080FFFF&
            Caption         =   "Send"
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
            Left            =   240
            Picture         =   "Form1.frx":2D82
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CALCULATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   5400
         TabIndex        =   47
         Top             =   360
         Width           =   7455
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   5520
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   720
            Width           =   1815
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            X1              =   0
            X2              =   7200
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C00000&
            Caption         =   "SE of Yield estimate GP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3360
            TabIndex        =   60
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C00000&
            Caption         =   "AVG  Yield /GP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label18 
            BackColor       =   &H0000FFFF&
            Caption         =   "SS-II"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label17 
            BackColor       =   &H0000FFFF&
            Caption         =   "SS-1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Productivity (Kg/ha)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   2
            Left            =   5400
            TabIndex        =   53
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Avg Mouza Production (Kg)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   1
            Left            =   2760
            TabIndex        =   52
            Top             =   360
            Width           =   2340
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "Avg Cropped Area/Mouza(ha)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   10
         Left            =   1800
         TabIndex        =   35
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   9
         Left            =   1800
         TabIndex        =   34
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   3720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   3480
         TabIndex        =   43
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3480
         TabIndex        =   41
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   30
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtArea 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtyield 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   375
         Left            =   5520
         TabIndex        =   63
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   1575
            Left            =   -360
            TabIndex        =   108
            Top             =   240
            Width           =   735
            Begin VB.TextBox txtY1 
               Height          =   285
               Left            =   6000
               TabIndex        =   114
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtY2 
               Height          =   285
               Left            =   6000
               TabIndex        =   113
               Top             =   1080
               Width           =   1215
            End
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   120
               TabIndex        =   112
               Top             =   600
               Width           =   615
            End
            Begin VB.TextBox txtprod1 
               Height          =   375
               Left            =   3840
               TabIndex        =   111
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtEarea1 
               Height          =   375
               Left            =   2040
               TabIndex        =   110
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox txtmiuza1 
               Height          =   285
               Left            =   960
               TabIndex        =   109
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "Y1"
               Height          =   255
               Index           =   0
               Left            =   6480
               TabIndex        =   120
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label9 
               Caption         =   "Y2"
               Height          =   255
               Index           =   1
               Left            =   6480
               TabIndex        =   119
               Top             =   840
               Width           =   375
            End
            Begin VB.Label Label19 
               Caption         =   "Total mouza"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   118
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label21 
               Caption         =   "Total Potato Production"
               Height          =   255
               Left            =   3720
               TabIndex        =   117
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label20 
               Caption         =   "Est Potato Area"
               Height          =   255
               Left            =   2280
               TabIndex        =   116
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label19 
               Caption         =   "Total of cuts"
               Height          =   255
               Index           =   0
               Left            =   1080
               TabIndex        =   115
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox Text20 
            Height          =   375
            Left            =   840
            TabIndex        =   107
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text19 
            Height          =   375
            Left            =   960
            TabIndex        =   106
            Text            =   "Text19"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox cs2 
            Height          =   285
            Left            =   720
            TabIndex        =   105
            Text            =   "Text13"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox cs1 
            Height          =   285
            Left            =   840
            TabIndex        =   104
            Text            =   "Text12"
            Top             =   600
            Width           =   495
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   975
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   615
            Begin VB.TextBox txtcut2 
               Height          =   285
               Left            =   960
               TabIndex        =   100
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtcut1 
               Height          =   285
               Left            =   0
               TabIndex        =   99
               Text            =   "0"
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label7 
               Caption         =   "Label4"
               Height          =   375
               Left            =   240
               TabIndex        =   101
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   135
            Left            =   960
            TabIndex        =   87
            Top             =   240
            Width           =   975
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   97
               Text            =   "no"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   96
               Text            =   "no"
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   3
               Left            =   120
               TabIndex        =   95
               Text            =   "no"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   94
               Text            =   "no"
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   5
               Left            =   120
               TabIndex        =   93
               Text            =   "no"
               Top             =   2160
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   6
               Left            =   120
               TabIndex        =   92
               Text            =   "no"
               Top             =   2520
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   91
               Text            =   "no"
               Top             =   2880
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   8
               Left            =   120
               TabIndex        =   90
               Text            =   "no"
               Top             =   3240
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   89
               Text            =   "no"
               Top             =   3600
               Width           =   735
            End
            Begin VB.TextBox txtno 
               Height          =   285
               Index           =   10
               Left            =   120
               TabIndex        =   88
               Text            =   "no"
               Top             =   3960
               Width           =   735
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "CAL Not need"
            Height          =   255
            Left            =   600
            TabIndex        =   71
            Top             =   240
            Width           =   735
            Begin VB.Frame Frame6 
               Caption         =   "PLT PRODUCT"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   240
               Width           =   1455
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   1
                  Left            =   120
                  TabIndex        =   86
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   2
                  Left            =   120
                  TabIndex        =   85
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   3
                  Left            =   120
                  TabIndex        =   84
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   4
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   5
                  Left            =   120
                  TabIndex        =   82
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   6
                  Left            =   120
                  TabIndex        =   81
                  Top             =   2040
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   7
                  Left            =   120
                  TabIndex        =   80
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   8
                  Left            =   120
                  TabIndex        =   79
                  Top             =   2760
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   9
                  Left            =   120
                  TabIndex        =   78
                  Top             =   3120
                  Width           =   1215
               End
               Begin VB.TextBox txtArea1 
                  Height          =   285
                  Index           =   10
                  Left            =   120
                  TabIndex        =   77
                  Top             =   3480
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "AVR AREA"
               Height          =   255
               Left            =   240
               TabIndex        =   74
               Top             =   480
               Width           =   1095
               Begin VB.TextBox TXTAVRAREA 
                  Height          =   375
                  Left            =   120
                  TabIndex        =   75
                  Text            =   "Text12"
                  Top             =   360
                  Width           =   855
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "AVR PL-PR"
               Height          =   255
               Left            =   240
               TabIndex        =   72
               Top             =   720
               Width           =   1095
               Begin VB.TextBox Text3 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1455
               End
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Can be Reprod/cal"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   615
            Begin VB.TextBox Text10 
               Height          =   375
               Left            =   3960
               TabIndex        =   67
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox Text8 
               Height          =   375
               Left            =   0
               TabIndex        =   66
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text7 
               Height          =   405
               Left            =   1800
               TabIndex        =   65
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label15 
               Caption         =   "EST. POT AREA"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label14 
               Caption         =   "EST NO. POTATO GR. PLT"
               Height          =   255
               Left            =   1680
               TabIndex        =   69
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label16 
               Caption         =   "TOT POT PRDU"
               Height          =   255
               Left            =   4080
               TabIndex        =   68
               Top             =   240
               Width           =   1335
            End
         End
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   3720
         Picture         =   "Form1.frx":364C
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "NO OF POTATO GROWING PLOTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3480
         TabIndex        =   42
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "PLT SURVEYED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "TOTAL PLOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "CROPPED AREA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1680
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "CCE YIELDS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   6720
      ScaleHeight     =   795
      ScaleWidth      =   4635
      TabIndex        =   134
      Top             =   2040
      Width           =   4695
      Begin VB.CommandButton Command5 
         BackColor       =   &H0000FF00&
         Caption         =   "INPUT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Picture         =   "Form1.frx":4516
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   0
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   240
         TabIndex        =   132
         Text            =   "Select Cuts No."
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "No. of Cuts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   135
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   6345
      TabIndex        =   126
      Top             =   2040
      Width           =   6375
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3960
         TabIndex        =   131
         Text            =   "Select Mouza No."
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2280
         TabIndex        =   129
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFC0FF&
         Caption         =   "New Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "Form1.frx":4DE0
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Mouza No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   130
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Sub Sampling"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   128
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      Picture         =   "Form1.frx":4F2A
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Edit"
      Height          =   615
      Left            =   12000
      Picture         =   "Form1.frx":57F4
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin EasyStat.LynxGrid LynxGrid 
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   7560
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   6588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeStyle      =   3
      AllowUserResizing=   4
      AutoSizeRow     =   -1  'True
      Editable        =   -1  'True
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "OK"
      Height          =   615
      Left            =   12000
      Picture         =   "Form1.frx":5B7E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   600
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox txtDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   11175
      End
      Begin VB.TextBox txtCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11175
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15345
      TabIndex        =   8
      Top             =   0
      Width           =   15375
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "POTATO SHEET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   6720
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   5640
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Text            =   "Select District"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Select year"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1035
      ScaleWidth      =   11355
      TabIndex        =   121
      Top             =   840
      Width           =   11415
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "G.P"
         Height          =   255
         Index           =   2
         Left            =   7560
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Block"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   124
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "District"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   123
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Year"
         Height          =   255
         Left            =   600
         TabIndex        =   122
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "info"
         Top             =   360
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      Caption         =   "NAME OF THE DISTRICT/BLOCK/GP AND YEAR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   4920
      TabIndex        =   139
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim m1 As Double 'TOTAL MOUZA
Dim a1 As Variant ' EST POTATO AREA
Dim p1 As Double 'TOTAL POTATO PRD
Dim S1 As Double
Dim A11 As Double 'avg cropped area/cal
Dim A2 As Double ' avg mouza prod
Dim y1 As Double  ' yield 1

Dim y2 As Double ' yield 2
Dim ym As Double ' average yield
Dim yr As Double

Dim ss As String
Dim mouza As Integer



Private Sub Combo2_Click()
Data1.Refresh
catquery = Combo2.Text

Data1.RecordSource = "SELECT * FROM info WHERE District =  '" & catquery & "' "

Data1.Refresh

Text1.Text = Data1.Recordset.Fields(0)
Data1.Recordset.Close
List1.Visible = True
Label2(1).Visible = True


End Sub

Private Sub Combo4_Click()
Frame5.Refresh
Frame5.Visible = False
For i = 1 To 10
txtyield(i).Visible = False
txtArea(i).Visible = False


Next i

TXTAVRAREA.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""

End Sub



Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Combo4.SetFocus
End If
End Sub

Private Sub Command1_Click()
Frame2.Visible = False
Combo1.Text = "Select year"
Combo2.Text = "Select District"
List1.Clear
List2.Clear
Label2(1).Visible = False
Label2(2).Visible = False


End Sub

Private Sub Command10_Click()

Unload Me
Me.Show
Me.Combo3.Text = "SS-I"


End Sub

Private Sub Command11_Click()

End Sub





Private Sub Command2_Click()
If Combo1.Text = "Select year" Or Combo1.Text = "" Then
MsgBox "SELECT THE YEAR"
Exit Sub
ElseIf Combo2.Text = "Select year" Or Combo2.Text = "" Then
MsgBox "SELECT THE DISTRICT"
Exit Sub
ElseIf List1.Text = "NO BLOCK" Or List1.Text = "" Then
MsgBox "SELECT THE BLOCK"
Exit Sub
ElseIf List2.Text = "  NO ITEM" Or List2.Text = "  NO G.P. on This Category" Or List2.Text = "" Then
MsgBox "SELECT THE GP"
Exit Sub
End If

Frame2.Visible = True
txtCaption.Text = "G.P.LEVEL ESTIMATION OF PRODUCTION AND YIELD RATE OF POTATO CROP FOR THE YEAR - " & Combo1.Text
txtDetails.Text = "DISTRICT - " & " " & Combo2.Text & "               " & "BLOCK - " & " " & List1.Text & "                    " & "G.P. - " & " " & List2.Text
List1.Visible = False
List2.Visible = False
Label2(1).Visible = False
Label2(2).Visible = False
Combo5.SetFocus

End Sub

Private Sub Command3_Click()
Unload Me
Me.Cls

End Sub

Private Sub Command4_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub Command5_Click()

If Not IsNumeric(Combo4.Text) = True Then
MsgBox "SELECT NO. OF CUTS"
Combo4.SetFocus
Exit Sub
End If


Text19.Text = Val(Text19.Text) + Val(Combo4.Text)
txtcut2 = Val(txtcut2.Text) + Val(Combo4.Text)
If Combo4.Text = "" Then
MsgBox "SELECT NO. OF CUTS"
Exit Sub
End If
Frame5.Refresh
Frame5.Visible = True
For i = 1 To Val(Combo4.Text)
txtyield(i).Visible = True
txtyield(i).Text = ""
txtyield(i).BackColor = vbWhite
txtArea(i).Visible = True
txtArea(i).Text = ""
txtArea(i).BackColor = vbWhite
txtArea1(i).Visible = True
txtArea1(i).Text = ""
txtno(i).Visible = True
txtno(i).BackColor = vbWhite
Next


For l = 1 To Val(Combo4.Text)
For n = Val(Text20.Text) + 1 To Val(Text19.Text)
txtno(l).Text = n
l = l + 1
Next
Next

Text4.BackColor = vbWhite
Text5.BackColor = vbWhite
Text6.BackColor = vbWhite
txtyield(1).SetFocus

End Sub

Private Sub Command6_Click()
If Text16.Text <> "" Then
MsgBox "You can not Calculate more than 2 Sub-Sample"
Exit Sub
End If

Text20.Text = Val(Text19.Text)


TXTAVRAREA.Text = (Val(txtArea(1)) + Val(txtArea(2)) + Val(txtArea(3)) + Val(txtArea(4)) + Val(txtArea(5)) + Val(txtArea(6)) + Val(txtArea(7)) + Val(txtArea(8)) + Val(txtArea(9)) + Val(txtArea(10))) / Val(Combo4.Text)

TXTAVRAREA.Text = Format(TXTAVRAREA.Text, "#.00")
txtArea1(1).Text = (Val(txtyield(1)) * Val(txtArea(1).Text)) / 0.5165
txtArea1(1).Text = Format(Val(txtArea1(1).Text), "#.00")

txtArea1(2).Text = (Val(txtyield(2)) * Val(txtArea(2).Text)) / 0.5165
txtArea1(2).Text = Format(Val(txtArea1(2).Text), "#.00")

txtArea1(3).Text = (Val(txtyield(3)) * Val(txtArea(3).Text)) / 0.5165
txtArea1(3).Text = Format(Val(txtArea1(3).Text), "#.00")

txtArea1(4).Text = (Val(txtyield(4)) * Val(txtArea(4).Text)) / 0.5165
txtArea1(4).Text = Format(Val(txtArea1(4).Text), "#.00")

txtArea1(5).Text = (Val(txtyield(5)) * Val(txtArea(5).Text)) / 0.5165
txtArea1(5).Text = Format(Val(txtArea1(5).Text), "#.00")

txtArea1(6).Text = (Val(txtyield(6)) * Val(txtArea(6).Text)) / 0.5165
txtArea1(6).Text = Format(Val(txtArea1(6).Text), "#.00")

txtArea1(7).Text = (Val(txtyield(7)) * Val(txtArea(7).Text)) / 0.5165
txtArea1(7).Text = Format(Val(txtArea1(7).Text), "#.00")

txtArea1(8).Text = (Val(txtyield(8)) * Val(txtArea(8).Text)) / 0.5165
txtArea1(8).Text = Format(Val(txtArea1(8).Text), "#.00")

txtArea1(9).Text = (Val(txtyield(9)) * Val(txtArea(9).Text)) / 0.5165
txtArea1(9).Text = Format(Val(txtArea1(9).Text), "#.00")

txtArea1(10).Text = (Val(txtyield(10)) * Val(txtArea(10).Text)) / 0.5165
txtArea1(10).Text = Format(Val(txtArea1(10).Text), "#.00")

Text3.Text = (Val(txtArea1(1)) + Val(txtArea1(2)) + Val(txtArea1(3)) + Val(txtArea1(4)) + Val(txtArea1(5)) + Val(txtArea1(6)) + Val(txtArea1(7)) + Val(txtArea1(8)) + Val(txtArea1(9)) + Val(txtArea1(10))) / Val(Combo4.Text)
Text3.Text = Format(Val(Text3.Text), "#.00")

If Text5.Text = "" Then
Text5.Text = "1"
End If



Text7.Text = Val(Text4.Text) * Val(Text6.Text) / Val(Text5.Text)
Text7.Text = Format(Val(Text7.Text), "#")

Text8.Text = Val(Text7.Text) * Val(TXTAVRAREA.Text)
Text8.Text = Format(Val(Text8.Text), "#.00")

Text10.Text = Val(Text7.Text) * Val(Text3.Text)
Text10.Text = Format(Text10.Text, "#.00")

'*********  SENDING THE CALCULATION TO FLEX GRID   ***************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim F As Long
F = Val(Combo4.Text)
cs1.Text = Val(cs2.Text) + 1
cs2.Text = Val(cs2.Text) + F


For i = 1 To F
LynxGrid.Redraw = False
LynxGrid.AddItem
LynxGrid.Redraw = True
Next
Label7.Caption = LynxGrid.ItemCount

Dim m  As Integer


For m = 1 To Val(Combo4.Text)
For j = Val(cs1.Text) To Val(cs2.Text)
LynxGrid.CellText(j + 1, 4) = txtyield(m).Text
LynxGrid.CellText(j + 1, 3) = txtno(m).Text
LynxGrid.CellText(j + 1, 5) = txtArea(m).Text
LynxGrid.CellText(j + 1, 7) = txtArea1(m).Text
m = m + 1



Next j
Next m

'''''for number



LynxGrid.CellText(Val(cs1.Text) + 1, 0) = Combo3.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 1) = Combo5.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 2) = Combo4.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 6) = TXTAVRAREA.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 8) = Text3.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 9) = Text4.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 10) = Text5.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 11) = Text6.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 12) = Text7.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 13) = TXTAVRAREA.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 14) = Text3.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 15) = Text8.Text
LynxGrid.CellText(Val(cs1.Text) + 1, 16) = Text10.Text
ss = Combo3.Text
Combo3.Text = ""

''''''''''''''''''''''  END OF FLEX GRID ''''''''''''''''''''''

'For i = 1 To k






Label7.Caption = LynxGrid.ItemCount

'calculation part '''''''''''''''''''''''''''''''''''''''

m1 = m1 + Val(Combo4.Text)
a1 = a1 + Val(Text8.Text)
p1 = p1 + Val(Text10.Text)
txtEarea1.Text = a1
txtprod1.Text = p1
txtmiuza1.Text = m1
mouza = mouza + 1

Text11.Text = mouza


'''initialised to zero value '''''''''''''''''''''''
For i = 1 To 10
txtArea(i).Text = ""
txtyield(i).Text = ""
Next

Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo4.SetFocus




End Sub

Private Sub Command7_Click()
If Text16.Text <> "" Then
MsgBox "You can not Calculate more than 2 Sub-Sample"
Exit Sub
End If
Dim lRow  As Long
 With LynxGrid
        .Redraw = False
         lRow = .AddItem()
        .CellText(lRow, 14) = "TOTAL"
        .CellFontBold(lRow, 14) = True
        .CellText(lRow, 15) = txtEarea1.Text
        .CellText(lRow, 16) = txtprod1.Text
        .Redraw = True
         End With
a1 = 0
p1 = 0
mouza = 0
m1 = 0

'''''''''total beneath cal ''''''''''

'''''''''''''''''''''''''''''''''''''
A11 = (Val((txtEarea1.Text) / 100) * 0.404686) / Val(Text11.Text)
A11 = Format(A11, "#.00")
If Not A11 <> 0 Then
 A11 = 1
 End If
A2 = (Val(txtprod1.Text)) / Val(Text11.Text)
A2 = Format(A2, "#.00")
y1 = Val(A2) / Val(A11)
y1 = Format(y1, "#.00")

''  for ss-I '''''''

If Len(Trim(Text13.Text)) <= 0 And Len(Trim(Text16.Text)) <= 0 Then
Text9.Text = A11
Text12.Text = A2
Text13.Text = y1

A11 = 0
A2 = 0
y1 = 0
ElseIf Len(Trim(Text13.Text)) <> 0 And Len(Trim(Text16.Text)) <= 0 Then
Text14.Text = A11
Text15.Text = A2
Text16.Text = y1

Text17.Text = (Val(Text13.Text) + Val(Text16.Text)) / 2
Text17.Text = Format(Text17.Text, "#.00")
Text18.Text = Abs(Val(Text13.Text) - Val(Text16.Text)) / 2
Text18.Text = Format(Text18.Text, "#.00")



A11 = 0
A2 = 0
y1 = 0
ElseIf Len(Trim(Text13.Text)) <> 0 And Len(Trim(Text16.Text)) <> 0 Then
MsgBox "You can not calculate - Please click New GP"
Exit Sub
A11 = 0
A2 = 0
y1 = 0
End If


Label7.Caption = LynxGrid.ItemCount
cs1.Text = Val(cs1.Text) + 1
cs2.Text = Val(cs2.Text) + 1

'''''for ss ii calculation'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'for area to initial stage
txtArea(1).Text = ""
txtArea(2).Text = ""
txtArea(3).Text = ""
txtArea(4).Text = ""
txtArea(5).Text = ""
txtArea(6).Text = ""
txtArea(7).Text = ""
txtArea(8).Text = ""
txtArea(9).Text = ""
txtArea(10).Text = ""

' for avg area to initial stage

TXTAVRAREA.Text = ""


txtArea1(1).Text = ""
txtArea1(2).Text = ""
txtArea1(3).Text = ""
txtArea1(4).Text = ""
txtArea1(5).Text = ""
txtArea1(6).Text = ""
txtArea1(7).Text = ""
txtArea1(8).Text = ""
txtArea1(9).Text = ""
txtArea1(10).Text = ""

Text3.Text = ""
txtEarea1.Text = ""

txtprod1.Text = ""

txtmiuza1.Text = ""
Text11.Text = "'"
For i = 1 To 10
txtArea(i).Text = ""
txtyield(i).Text = ""
Next
Command9.SetFocus
End Sub

Private Sub Command8_Click()
a = 0
Dim i As Long
Dim n As Long
Dim xlapp As Excel.Application
Dim xlwb As Excel.Workbook
Dim xlws As Excel.Worksheet
On Error Resume Next
Set xlapp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
Set xlapp = New Excel.Application
Err.Clear
End If

Set xlwb = xlapp.Workbooks.Add
Set xlws = xlwb.ActiveSheet

xlapp.Visible = True
xlwb.ActiveSheet.Cells(1, 1).Value = txtCaption.Text
xlwb.ActiveSheet.Cells(2, 1).Value = txtDetails


'For i = 1 To Val(Label5.Caption) - 1
   ' OrderGrid.Row = i
   ' For n = 0 To 8
      '  OrderGrid.Col = n
   '    ' xlwb.ActiveSheet.Cells(i + 5, n + 1).Value = OrderGrid.Text
    'Next
'Next



For i = 0 To (Val(cs2.Text) + 1)
   LynxGrid.Row = i
    For n = 0 To 16
       LynxGrid.Col = n
        xlwb.ActiveSheet.Cells(i + 6, n + 1).Value = LynxGrid.CellText(LynxGrid.Row, LynxGrid.Col)
    Next
Next

''***********  CALCULATION FOR SS-I **************''


xlwb.ActiveSheet.Cells(21, 1).Value = "Calculation for " & Label17.Caption


xlwb.ActiveSheet.Cells(22, 1).Value = "Average Cropped Area per mouza(ha)=A={(C0L 15/100)X 0.404686}/No of Mouzas =  " & Text9.Text

xlwb.ActiveSheet.Cells(23, 1).Value = "Average Mouza Production(kg) =P= Col 16/No of Mouzas =  " & Text12.Text

xlwb.ActiveSheet.Cells(24, 1).Value = "Productivity (Kg/ha)P/A =  " & Text13.Text

''***********  CALCULATION FOR SS-II **************''


xlwb.ActiveSheet.Cells(26, 1).Value = "Calculation for " & Label18.Caption


xlwb.ActiveSheet.Cells(27, 1).Value = "Average Cropped Area per mouza(ha)=A={(C0L 15/100)X 0.404686}/No of Mouzas =  " & Text14.Text

xlwb.ActiveSheet.Cells(28, 1).Value = "Average Mouza Production(kg) =P= Col 16/No of Mouzas =  " & Text15.Text

xlwb.ActiveSheet.Cells(29, 1).Value = "Productivity (Kg/ha)P/A =  " & Text16.Text

''***********  CALCULATION FOR SUMMARY **************''

xlwb.ActiveSheet.Cells(31, 1).Value = "SUMMARY"
xlwb.ActiveSheet.Cells(32, 1).Value = "Average Yield for the G.P = (Y1 + Y2)/2 =  " & Text17.Text
xlwb.ActiveSheet.Cells(33, 1).Value = "S.E  of Yield Estimate of the G.P = abs(Y1-Y2) =  " & Text18.Text



'''macro
  Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("6:6").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Selection.ColumnWidth = 4.86
    Range("B6").Select
    ActiveWindow.SmallScroll ToRight:=3
    Range("B6:R20").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("E6").Select
    Columns("H:H").ColumnWidth = 6.29
    Columns("I:I").ColumnWidth = 9.71
    ActiveWindow.SmallScroll Down:=-3
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(1)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 180
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLegal
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    Selection.Font.Bold = True
    Rows("2:2").Select
    Selection.Font.Bold = True
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-6
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = ""
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 180
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLegal
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
    ActiveWindow.SmallScroll Down:=9
    Rows("21:21").Select
    Selection.Font.Bold = True
    Rows("26:26").Select
    Selection.Font.Bold = True
    Rows("31:31").Select
    Selection.Font.Bold = True
    Range("C36").Select
    ActiveWindow.SmallScroll Down:=-12
    
      ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("B6").Select
    ActiveWindow.SmallScroll ToRight:=4
    Range("B6:R6").Select
    ActiveWindow.SmallScroll Down:=6
    Range("B6:R20").Select
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("L22").Select
    
    Columns("M:M").ColumnWidth = 7.29
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("E:E").ColumnWidth = 7
    Columns("D:D").ColumnWidth = 6
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    Columns("R:R").ColumnWidth = 12.29
    Columns("R:R").ColumnWidth = 12.71
    Columns("Q:Q").ColumnWidth = 9.57
    Columns("O:O").ColumnWidth = 7.57
    Rows("7:7").Select
    Range("D7").Activate
    Selection.Font.Bold = True
    Range("M10").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Columns("B:B").Select
    Selection.Font.Bold = False
    Rows("1:1").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Rows("2:2").Select
    Selection.Font.Bold = False
    Selection.Font.Bold = True
    Range("B3:B20").Select
    Selection.Font.Bold = True
    Range("C8").Select
    Selection.Font.Bold = True
    Columns("E:E").ColumnWidth = 5.86
    Range("K24").Select

'''' END  SENDING GRID TO ExCEL'''''






Set xlapp = Nothing
Set xlwb = Nothing
Set xlws = Nothing


End Sub

Private Sub Command9_Click()
Me.Combo3.Text = "SS-II"

End Sub

Private Sub Form_Activate()
Me.Combo3.Text = "SS-I"
'''lynx grid
Dim lRow As Long

    With LynxGrid
        .Redraw = False
        
        .EditTrigger = lgEnterKey
        .FullRowSelect = False
        
        .FocusRectColor = &HFF0000
        .FocusRectMode = lgRow
        .FocusRectStyle = lgFRHeavy
        
        .RowHeightMin = 315
        
       
        
        .AddColumn "(0)", 450, , , , lgAlignCenterTop, True
        .AddColumn "(1)", 550, , , , lgAlignCenterTop, True
        .AddColumn "(2)", 350, , , , lgAlignCenterTop, True
        .AddColumn "(3)", 350, , , , lgAlignCenterTop, True
         .AddColumn "(4)", 920, , , , lgAlignCenterTop, True
          .AddColumn "(5)", 750, , , , lgAlignCenterTop, True
           .AddColumn "(6)", 750, , , , lgAlignCenterTop, True
            .AddColumn "(7)", 1550, , , , lgAlignCenterTop, True
             .AddColumn "(8)", 1550, , , , lgAlignCenterTop, True
              .AddColumn "(9)", 950, , , , lgAlignCenterTop, True
               .AddColumn "(10)", 650, , , , lgAlignCenterTop, True
                .AddColumn "(11)", 550, , , , lgAlignCenterTop, True
                 .AddColumn "(12)", 850, , , , lgAlignCenterTop, True
                  .AddColumn "(13)", 550, , , , lgAlignCenterTop, True
                   .AddColumn "(14)", 950, , , , lgAlignCenterTop, True
                    .AddColumn "(15)", 1550, , , , lgAlignCenterTop, True
                     .AddColumn "(16)", 1850, , , , lgAlignCenterTop, True
        
        .BindControl 1, txtEdit
        .Redraw = True
    End With
    
    '''fill up the blanks'''
    
  


Dim lRow1  As Long

 With LynxGrid
        .Redraw = False
         lRow1 = .AddItem()
        .CellText(lRow1, 0) = "SS"
        .CellFontBold(lRow1, 0) = True
        .CellText(lRow1, 1) = "JL No."
        .CellText(lRow1, 2) = "No. of Cuts"
        .CellText(lRow1, 3) = "Cuts Sl No"
        .CellText(lRow1, 4) = "C.C.E Yield (Kg)"
        .CellText(lRow1, 5) = "Cropped Area (decimal) in the plots"
        .CellText(lRow1, 6) = "Average Cropped Area"
        .CellText(lRow1, 7) = "Plot Production (Kg.)Col 5X4/0.5165"
        .CellText(lRow1, 8) = "Average Plot Production(kg)"
        .CellText(lRow1, 9) = "Total No. of plots in the mouza"
        .CellText(lRow1, 10) = "No. of plots suryeyed"
        .CellText(lRow1, 11) = "No. of potato growing plot found"
        .CellText(lRow1, 12) = "Est. No. of potato growing plot Col 9 X 11/10"
        .CellText(lRow1, 13) = "Average Cropped Area"
        .CellText(lRow1, 14) = "Average production(kg)per plot of mouza"
        .CellText(lRow1, 15) = "Est.potato area(decimal)in selected mouza  Col 12X13"
        .CellText(lRow1, 16) = "Total potato production in selected mouzas  Col 12X14"
        .Redraw = True
         End With
         
         
         
         LynxGrid.Redraw = False
         LynxGrid.AddItem
            For i = 0 To 16
           LynxGrid.CellText(1, i) = "[" & i & "]"
           Next
            LynxGrid.Redraw = True
            
         
 
        


Label7.Caption = LynxGrid.ItemCount
End Sub

Private Sub Form_Load()

Data1.DatabaseName = App.Path & "/database/db2.mdb"
Data3.DatabaseName = App.Path & "/database/db2.mdb"
Data2.DatabaseName = App.Path & "/database/db2.mdb"

Combo1.AddItem "2006 - 07"
Combo1.AddItem "2007 - 08"
Combo1.AddItem "2008 - 09"
Combo1.AddItem "2009 - 10"
Combo1.AddItem "2010 - 11"

Combo3.AddItem "SS-I"
Combo3.AddItem "SS-II"
 
 For j = 1 To 1000
 Combo5.AddItem j
 Next
 
For t = 1 To 10

Combo4.AddItem t
Next t

Data1.Refresh
Do While Not Data1.Recordset.EOF
Combo2.AddItem (Data1.Recordset.Fields(1))
Data1.Recordset.MoveNext
Loop
Data1.Recordset.Close



    
End Sub

Private Sub List1_Click()
Data2.Refresh
catquery = List1.Text

Data2.RecordSource = "SELECT * FROM block WHERE Block =  '" & catquery & "' "

Data2.Refresh

Text2.Text = Data2.Recordset.Fields(1)
Data2.Recordset.Close
List2.Visible = True

Label2(2).Visible = True

End Sub

Private Sub List2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command2.SetFocus
End If

End Sub

Private Sub LynxGrid_CellImageClick(ByVal Row As Long, ByVal Col As Long)
 If Col = 1 Then
        With LynxGrid
            If .RowHeight(Row) = .RowHeightMin Then
                .RowHeight(Row) = -1
            Else
                .RowHeight(Row) = .RowHeightMin
            End If
        End With
    End If
    
End Sub

Private Sub LynxGrid_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
        txtEdit.Text = LynxGrid.CellText(Row, Col)
        End If
End Sub

Private Sub LynxGrid_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
If Col = 1 Then
        NewValue = txtEdit.Text
    End If
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
List1.Clear
List1.AddItem "NO BLOCK"
Exit Sub
End If
List1.Clear
Data2.Refresh
Data2.RecordSource = "SELECT * FROM block WHERE Sl =  '" & Text1.Text & "' "
Data2.Refresh
Do While Not Data2.Recordset.EOF
List1.AddItem (Data2.Recordset.Fields(2))
Data2.Recordset.MoveNext
Loop
Data2.Recordset.Close




End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
List2.Clear
List2.AddItem "  NO ITEM"
Exit Sub
End If
List2.Clear
Data3.Refresh
Data3.RecordSource = "SELECT * FROM GP WHERE Bsl =  '" & Text2.Text & "' "
Data3.Refresh

If Data3.Recordset.RecordCount <= 0 Then
List2.AddItem "  NO G.P. on This Category"
End If
Do While Not Data3.Recordset.EOF
List2.AddItem (Data3.Recordset.Fields(1))
Data3.Recordset.MoveNext
Loop
Data3.Recordset.Close
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = vbYellow
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = vbYellow
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = vbYellow
End Sub

Private Sub txtArea_GotFocus(Index As Integer)
For m = 1 To Val(Combo4.Text)

txtArea(m).BackColor = vbYellow

Next
End Sub

Private Sub txtyield_GotFocus(Index As Integer)
For m = 1 To Val(Combo4.Text)

txtyield(m).BackColor = vbYellow

Next
End Sub

