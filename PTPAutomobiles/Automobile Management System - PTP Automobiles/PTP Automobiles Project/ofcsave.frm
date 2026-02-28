VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ofcsave 
   Caption         =   "Take Order"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18390
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCustContact 
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
      Left            =   10440
      TabIndex        =   7
      Top             =   4800
      Width           =   3255
   End
   Begin VB.TextBox txtCustName 
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
      Left            =   10440
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox txtOId 
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
      Left            =   10440
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox txtTotalAmount 
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
      Left            =   10440
      TabIndex        =   14
      Top             =   6840
      Width           =   3255
   End
   Begin VB.TextBox txtCustAddress 
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
      Left            =   10440
      TabIndex        =   8
      Top             =   5280
      Width           =   3255
   End
   Begin VB.TextBox txtVModelNo 
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
      Left            =   10440
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtVName 
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
      Left            =   10440
      TabIndex        =   3
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtCustId 
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
      Left            =   10440
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.TextBox txtVCompany 
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
      Left            =   10440
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtDiscount 
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
      Left            =   10440
      TabIndex        =   15
      Top             =   7320
      Width           =   3255
   End
   Begin VB.ComboBox cmbOTakenBy 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "ofcsave.frx":0000
      Left            =   10440
      List            =   "ofcsave.frx":0002
      TabIndex        =   13
      Top             =   6240
      Width           =   3495
   End
   Begin VB.OptionButton p1 
      Caption         =   "Cash"
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
      Left            =   10440
      TabIndex        =   9
      Top             =   5760
      Width           =   1935
   End
   Begin VB.OptionButton p3 
      Caption         =   "Net Banking"
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
      Left            =   14520
      TabIndex        =   11
      Top             =   5760
      Width           =   2295
   End
   Begin VB.OptionButton p2 
      Caption         =   "Cheque"
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
      Left            =   12480
      TabIndex        =   10
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
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
      Left            =   8040
      TabIndex        =   17
      Top             =   8640
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   10800
      TabIndex        =   20
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
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
      Left            =   8040
      TabIndex        =   19
      Top             =   9720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   10800
      TabIndex        =   18
      Top             =   8640
      Width           =   2415
   End
   Begin VB.OptionButton p4 
      Caption         =   "Card"
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
      Left            =   16920
      TabIndex        =   12
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtNetAmount 
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
      Left            =   10440
      TabIndex        =   16
      Top             =   7800
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
      TabIndex        =   21
      Top             =   0
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtODate 
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110231553
      CurrentDate     =   44231
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "Order_Date"
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
      Left            =   7920
      TabIndex        =   35
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Customer Contact"
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
      Left            =   7920
      TabIndex        =   34
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Customer Name"
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
      Left            =   7920
      TabIndex        =   33
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Order_Id"
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
      Left            =   7920
      TabIndex        =   32
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Vehicle Model No."
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
      Left            =   7920
      TabIndex        =   31
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Vehicle Name"
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
      Left            =   7920
      TabIndex        =   30
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Customer_Id"
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
      Left            =   7920
      TabIndex        =   29
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Vehicle Company"
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
      Left            =   7920
      TabIndex        =   28
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Discount"
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
      Left            =   7920
      TabIndex        =   27
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Total Amount"
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
      Left            =   7920
      TabIndex        =   26
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Payment Mode"
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
      Left            =   7920
      TabIndex        =   25
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Order Taken By"
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
      Left            =   7920
      TabIndex        =   24
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Net Amount"
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
      Left            =   7920
      TabIndex        =   23
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Customer Address"
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
      Left            =   7920
      TabIndex        =   22
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Image i1 
      Height          =   10935
      Left            =   0
      Picture         =   "ofcsave.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "ofcsave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim temp As Double
    temp = Val(txtTotalAmount.Text * (txtDiscount.Text / 100))
    txtNetAmount.Text = Val(txtTotalAmount.Text - temp)
End Sub
Private Sub cmdExit_Click()
    Load home
    home.Show
    ofcsave.Hide
    Unload ofcsave
End Sub
Private Sub cmdPrevious_Click()
    Load ofc
    ofc.Show
    ofcsave.Hide
    Unload ofcsave
End Sub
Private Sub cmdSave_Click()
    Dim ans As String
    ans = MsgBox("Do you want to Save this Record, then click Yes", vbYesNo)
    
    If ans = vbYes Then
    recset.AddNew
    Call updateRecord2
    recset.Update
    MsgBox "Your record is Saved..."
    Call cmdClear_Click
    End If

    If ans = vbNo Then
    recset.CancelUpdate
    MsgBox "Your saving is cancelled!!!"
    End If
End Sub

Private Sub autoGenerateId()
Call closeDB
Call dbconnection2
    If recset.BOF = True Or recset.EOF = True Then
        txtOId.Text = 1000
        txtCustId.Text = 1000
    Else
        recset.MoveLast
        txtOId.Text = recset.Fields("Order_Id").Value + 1
        txtCustId.Text = recset.Fields("Customer_Id").Value + 1
    End If
End Sub
Private Sub updateRecord2()
    recset.Fields("Order_Id").Value = Val(txtOId.Text)
    recset.Fields("Order_Date").Value = dtODate.Value
    recset.Fields("Vehicle_Company").Value = txtVCompany.Text
    recset.Fields("Vehicle_Name").Value = txtVName.Text
    recset.Fields("Vehicle_Model_No").Value = txtVModelNo.Text
    recset.Fields("Customer_Id").Value = Val(txtCustId.Text)
    recset.Fields("Customer_Name").Value = txtCustName.Text
    recset.Fields("Customer_Contact").Value = txtCustContact.Text
    recset.Fields("Customer_Address").Value = txtCustAddress.Text
    If p1.Value = True Then
        recset.Fields("Payment_Mode").Value = p1.Caption
    End If
    If p2.Value = True Then
        recset.Fields("Payment_Mode").Value = p2.Caption
    End If
    If p3.Value = True Then
        recset.Fields("Payment_Mode").Value = p3.Caption
    End If
    If p4.Value = True Then
        recset.Fields("Payment_Mode").Value = p4.Caption
    End If
    recset.Fields("Order_Taken_By").Value = cmbOTakenBy.Text
    recset.Fields("Total_Amount").Value = Val(txtTotalAmount.Text)
    recset.Fields("Discount").Value = Val(txtDiscount.Text)
    recset.Fields("Net_Amount").Value = Val(txtNetAmount)
End Sub
Private Sub cmdClear_Click()
    txtOId.Text = ""
    dtODate.Value = Date
    txtVCompany.Text = ""
    txtVName.Text = ""
    txtVModelNo.Text = ""
    txtCustId.Text = ""
    txtCustName.Text = ""
    txtCustContact.Text = ""
    txtCustAddress.Text = ""
    p1.Value = False
    p2.Value = False
    p3.Value = False
    p4.Value = False
    cmbOTakenBy.Text = ""
    txtTotalAmount.Text = ""
    txtDiscount.Text = ""
    txtNetAmount.Text = ""
    Call autoGenerateId
End Sub
Private Sub addingEmployee()
    Call dbconnection5
    recset.MoveFirst
    Do Until recset.EOF = True
        cmbOTakenBy.AddItem recset.Fields("Name").Value
        recset.MoveNext
    Loop
    recset.MoveFirst
    Call closeDB
End Sub

Private Sub Form_Load()
    Call addingEmployee
    Call dbconnection2
    Call autoGenerateId
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call closeDB
End Sub
Private Sub Form_Resize()
i1.Width = Me.ScaleWidth
i1.Height = Me.ScaleHeight
End Sub
