VERSION 5.00
Begin VB.Form otm 
   Caption         =   "Choose Record"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "otm.frx":0000
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd4 
      BackColor       =   &H80000010&
      Caption         =   "Remove Order"
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
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9000
      Width           =   4215
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H80000010&
      Caption         =   "Search Order"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      Width           =   4215
   End
   Begin VB.CommandButton cmd2 
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
      Height          =   1335
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   4215
   End
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
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "otm.frx":59559
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "otm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Load home
home.Show
otm.Hide
Unload otm
End Sub
Private Sub cmd2_Click()
Load otmsave
otmsave.Show
otm.Hide
Unload otm
End Sub
Private Sub cmd3_Click()
Load otmsearch
otmsearch.Show
otm.Hide
Unload otm
End Sub
Private Sub cmd4_Click()
Load otmdelete
otmdelete.Show
otm.Hide
Unload otm
End Sub
Private Sub Form_Resize()
i1.Width = Me.ScaleWidth
i1.Height = Me.ScaleHeight
End Sub
