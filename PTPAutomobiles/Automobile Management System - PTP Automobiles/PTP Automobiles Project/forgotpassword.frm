VERSION 5.00
Begin VB.Form forgotpassword 
   Caption         =   "Forgot Password"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   15390
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   3600
      Width           =   3615
   End
   Begin VB.CommandButton cmdback 
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
      TabIndex        =   5
      Top             =   0
      Width           =   2415
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
      TabIndex        =   0
      Top             =   1560
      Width           =   3255
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
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox txtNewPass 
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
      TabIndex        =   2
      Top             =   2520
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
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label1 
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
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label4 
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
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "New Password"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label6 
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
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "forgotpassword.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20175
   End
End
Attribute VB_Name = "forgotpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Load login
login.Show
forgotpassword.Hide
Unload forgotpassword
End Sub
Private Sub cmdChangePassword_Click()
Call closeDB
    Dim ans As String
    dbconn.Open "provider=MSDASQL;driver={Mysql odbc 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
    recset.Open "select * from Employee where Username = '" & txtUsername.Text & "' And MobileNo = '" & txtMobileNo.Text & "'", dbconn, adOpenDynamic, adLockBatchOptimistic
    If recset.EOF = True Then
        MsgBox "Username or Mobile Number is  Invalid!!!"
        Call closeDB
        Exit Sub
    Else
        ans = MsgBox("Do you want to Change Password ?, then click Yes", vbYesNo)
        If ans = vbYes Then
            If txtNewPass.Text = txtConPass.Text Then
             'If StrComp(Trim(txtNewPass.Text), Trim(txtConPass.Text)) = 0 Then
                Call changePass
                recset.Update
                MsgBox "Your Password is Successfully Changed..."
                Call ClearAll
                Call closeDB
                Exit Sub
             'End If
            Else
                MsgBox "Your Confirm Password is not matched with New Password!!!"
                Call closeDB
                Exit Sub
           End If
        End If
        If ans = vbNo Then
            recset.CancelUpdate
            MsgBox "Password Not Updated!"
            Call closeDB
        End If
    End If
End Sub
Private Sub changePass()
Dim temp As String
temp = txtConPass.Text
recset.Fields("Password").Value = Trim(temp)
End Sub

Private Sub ClearAll()
    txtUsername.Text = ""
    txtMobileNo.Text = ""
    txtNewPass.Text = ""
    txtConPass.Text = ""
End Sub

Private Sub Form_Load()
    Call dbconnection5
End Sub

Private Sub Form_Resize()
i1.Width = Me.ScaleWidth
i1.Height = Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call closeDB
End Sub
