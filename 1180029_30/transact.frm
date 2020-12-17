VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "back to selection"
      Height          =   1695
      Left            =   12240
      TabIndex        =   9
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdtransact 
      Caption         =   "transact"
      Height          =   1815
      Left            =   9120
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtnewbalance 
      Height          =   855
      Left            =   8400
      TabIndex        =   6
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txttransact 
      Height          =   855
      Left            =   2760
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.OptionButton optwithdraw 
      Caption         =   "withdraw"
      Height          =   1215
      Left            =   5280
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.OptionButton optdeposit 
      Caption         =   "deposit"
      Height          =   1095
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Txtcurbalance 
      Height          =   735
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Txtaccno 
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox cmbcustid 
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label newbalancelbl 
      Caption         =   "new balance"
      Height          =   735
      Left            =   6120
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label transactlbl 
      Caption         =   "transaction amount"
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label curballbl 
      Caption         =   "curent balance"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label accnolbl 
      Caption         =   "account no"
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label custidlbl 
      Caption         =   "customer id"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Dim rs As ADODB.Recordset

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\banking.mdb;Persist Security Info=False"
con.Open
rs.Open " select * from custtbl", con, 1, 2
While Not (rs.EOF)
cmbcustid.AddItem (rs!customerid)
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub cmbcustid_Click()
rs.Open "select * from custtbl where customerid=" & cmbcustid.Text, con, 1, 2
Txtaccno.Text = rs!accountno
Txtcurbalance.Text = rs!balance
rs.Close
End Sub

Private Sub optdeposit_Click()
txtnewbalance.Text = Val(Txtcurbalance.Text) + Val(txttransact.Text)
End Sub

Private Sub optwithdraw_Click()
txtnewbalance.Text = Val(Txtcurbalance.Text) - Val(txttransact.Text)
End Sub

Private Sub cmdtransact_Click()
rs.Open "select * from custtbl where customerid=" & cmbcustid.Text, con, 1, 2
rs!balance = Val(txtnewbalance.Text)
rs.Update
rs.Close
End Sub


Private Sub cmdexit_Click()
Unload Me
Form2.Show
End Sub
