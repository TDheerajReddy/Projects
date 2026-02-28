VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form singup 
   Caption         =   "Create New Account"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17730
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   17730
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dtDOB 
      Height          =   375
      Left            =   10200
      TabIndex        =   18
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110034945
      CurrentDate     =   44300
   End
   Begin VB.TextBox txtConPass 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   10200
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   4
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtMobileNo 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CommandButton cmdSignUp 
      Caption         =   "Sing Up"
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
      Left            =   8520
      TabIndex        =   7
      Top             =   6240
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   0
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrevious 
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
      TabIndex        =   8
      Top             =   0
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtDOJ 
      Height          =   375
      Left            =   10200
      TabIndex        =   19
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110034945
      CurrentDate     =   44300
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Mobile Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Date Of Birth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Date Of Joining"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "singup.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "singup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrevious_Click()
    Load login
    login.Show
    singup.Hide
    Unload singup
End Sub

Private Sub cmdSignUp_Click()
Call dbconnection5
Dim ans As String
    ans = MsgBox("Do you want to Create Account ?, then click Yes", vbYesNo)
    If ans = vbYes Then
        Dim temp1, temp2 As String
        temp1 = txtPass.Text
        temp2 = txtConPass.Text
            If temp1 = temp2 Then
                recset.AddNew
                Call updateRecord5
                recset.Update
                MsgBox "Your Account is Created..."
                Call closeDB
                Exit Sub
            Else
                MsgBox "Confirm Password is not same from Original Password"
                Call closeDB
                Exit Sub
            End If
        Call closeDB
        Exit Sub
    End If

    If ans = vbNo Then
        recset.CancelUpdate
        MsgBox "Your Account creation is cancelled!!!"
    End If
Call closeDB
End Sub
Private Sub updateRecord5()
    recset.Fields("Name").Value = txtName.Text
    recset.Fields("Address").Value = txtAddress.Text
    recset.Fields("MobileNo").Value = txtMobileNo.Text
    recset.Fields("Email").Value = txtEmail.Text
    recset.Fields("DOB").Value = dtDOB.Value
    recset.Fields("DOJ").Value = dtDOJ.Value
    recset.Fields("Username").Value = txtUsername.Text
    recset.Fields("Password").Value = txtPass.Text
End Sub
Private Sub txtAddress_Validate(Cancel As Boolean)
If Len(txtAddress.Text) <= 50 Then
    Cancel = False
    Exit Sub
Else
    MsgBox "Address should be atmost 50 Characters!!!"
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub txtEmail_Validate(Cancel As Boolean)
If Len(txtAddress.Text) <= 50 Then
    Cancel = False
    Exit Sub
Else
    MsgBox "Address should be atmost 50 Characters!!!"
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub Form_Resize()
    i1.Width = Me.ScaleWidth
    i1.Height = Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call closeDB
End Sub
