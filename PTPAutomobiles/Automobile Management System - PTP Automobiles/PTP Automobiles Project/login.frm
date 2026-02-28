VERSION 5.00
Begin VB.Form login 
   Caption         =   "Please Login"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   17925
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRemAccount 
      Caption         =   "Remove Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdNewAccount 
      Caption         =   "Create  New Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton cmdForgotPass 
      Caption         =   "Forgotten Password?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   10320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
      Width           =   3015
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   10320
      TabIndex        =   0
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7680
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForgotPass_Click()
    Load forgotpassword
    forgotpassword.Show
    login.Hide
    Unload login
End Sub

Private Sub cmdLogin_Click()
Call closeDB
    Call dbconnection5
    Dim temp1, temp2 As String
    temp1 = txtUsername.Text
    temp2 = txtPassword.Text

    Set recset = Nothing
    recset.Open "select * from Employee where Username = '" & temp1 & "' and Password = '" & temp2 & "'", dbconn, adOpenDynamic, adLockOptimistic
    If recset.EOF = True Then
        MsgBox "Incorrect Username or Password, Try Again!"
        Call closeDB
        Exit Sub
    Else
        Load home
        home.Show
        login.Hide
        Unload login
        Call closeDB
        Exit Sub
    End If
    Call closeDB
End Sub
Private Sub cmdNewAccount_Click()
    Load singup
    singup.Show
    login.Hide
    Unload login
End Sub

Private Sub cmdRemAccount_Click()
Call closeDB
dbconn.Open "provider=MSDASQL;driver={Mysql odbc 3.51 Driver};database=PTPAutomobiles;server=localhost;user=root;password=123"
recset.Open "select * from Employee where Username = '" & txtUsername.Text & "' And Password = '" & txtPassword.Text & "'", dbconn, adOpenDynamic, adLockOptimistic
    If recset.EOF = True Then
        MsgBox "RECORD NOT FOUND!!!"
        Call closeDB
        Exit Sub
    Else
        ans = MsgBox("Do you want to delete this Record, then click Yes", vbYesNo)
        If ans = vbYes Then
            recset.Delete
            MsgBox "RECORD DELETED..."
            Call ClearAll
            Call closeDB
            Exit Sub
        End If
        If ans = vbNo Then
            recset.Cancel
            MsgBox "Your Deletion is Cancelled!!!"
            Call closeDB
            Exit Sub
        End If
    End If
End Sub

Private Sub ClearAll()
    txtUsername.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub Form_Resize()
    i1.Width = Me.ScaleWidth
    i1.Height = Me.ScaleHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call closeDB
End Sub

