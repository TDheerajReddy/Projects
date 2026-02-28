VERSION 5.00
Begin VB.Form spareparts 
   Caption         =   "Choose  Parts Order"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd1 
      Caption         =   "<Previous"
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
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H80000010&
      Caption         =   "Save Parts Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   4215
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H80000010&
      Caption         =   "Search Parts Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   4215
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H80000010&
      Caption         =   "Remove Parts Order"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   4215
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "spareparts.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "spareparts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Load home
home.Show
spareparts.Hide
Unload spareparts
End Sub
Private Sub cmd2_Click()
Load sparepartssave
sparepartssave.Show
spareparts.Hide
Unload spareparts
End Sub
Private Sub cmd3_Click()
Load sparepartssearch
sparepartssearch.Show
spareparts.Hide
Unload spareparts
End Sub
Private Sub cmd4_Click()
Load sparepartsdelete
sparepartsdelete.Show
spareparts.Hide
Unload spareparts
End Sub

Private Sub Form_Resize()
i1.Width = Me.ScaleWidth
i1.Height = Me.ScaleHeight
End Sub
