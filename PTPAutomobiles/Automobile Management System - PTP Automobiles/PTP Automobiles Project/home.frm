VERSION 5.00
Begin VB.Form home 
   Caption         =   "What Do You Have To Do"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdparts 
      BackColor       =   &H80000010&
      Caption         =   "Spare Parts"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9480
      Width           =   4095
   End
   Begin VB.CommandButton cmdservice 
      BackColor       =   &H80000010&
      Caption         =   "Servicing "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9480
      Width           =   3975
   End
   Begin VB.CommandButton cmdtake 
      BackColor       =   &H80000010&
      Caption         =   "Take Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   4095
   End
   Begin VB.CommandButton cmdgive 
      BackColor       =   &H80000010&
      Caption         =   "Give Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Option"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   8040
      Width           =   5055
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "home.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Load welcome
    welcome.Show
    home.Hide
    Unload home
End Sub

Private Sub cmdgive_Click()
    Load otm
    otm.Show
    home.Hide
    Unload home
End Sub

Private Sub cmdparts_Click()
    Load spareparts
    spareparts.Show
    home.Hide
    Unload home
End Sub

Private Sub cmdservice_Click()
    Load service
    service.Show
    home.Hide
    Unload home
End Sub

Private Sub cmdtake_Click()
    Load ofc
    ofc.Show
    home.Hide
    Unload home
End Sub

Private Sub Form_Resize()
i1.Height = Me.ScaleHeight
i1.Width = Me.ScaleWidth
End Sub
