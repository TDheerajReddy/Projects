VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form process 
   Caption         =   "Process"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer T1 
      Interval        =   100
      Left            =   2760
      Top             =   10200
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   10200
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label L2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   1
      Top             =   10200
      Width           =   735
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "Process Form.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
i1.Height = Me.ScaleHeight
i1.Width = Me.ScaleWidth
End Sub

Private Sub T1_Timer()
T1.Interval = Rnd * 100 + 10
pb1.Value = pb1.Value + 5
L1.Caption = pb1.Value & "%"
If L1.Caption = 100 & "%" Then
Load welcome
welcome.Show
process.Hide
Unload process
End If
End Sub
