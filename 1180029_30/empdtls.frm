VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtnetsal 
      Height          =   975
      Left            =   12240
      TabIndex        =   24
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox Txttax 
      Height          =   855
      Left            =   12240
      TabIndex        =   23
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Txtgrosssal 
      Height          =   495
      Left            =   6000
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Txtothrs 
      Height          =   495
      Left            =   6120
      TabIndex        =   19
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Txtwkhrs 
      Height          =   615
      Left            =   5880
      TabIndex        =   16
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "cancel"
      Height          =   735
      Left            =   7800
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdcal 
      Caption         =   "calculate salary"
      Height          =   735
      Left            =   5280
      TabIndex        =   6
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox Txtotrate 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Txtrph 
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Txtempage 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Txtempgender 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Txtempname 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ComboBox cmbempid 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "netsal"
      Height          =   735
      Left            =   9120
      TabIndex        =   22
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "taxded"
      Height          =   495
      Left            =   9240
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "grosssalary"
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "othrs"
      Height          =   615
      Left            =   4080
      TabIndex        =   17
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "hrsw"
      Height          =   615
      Left            =   4080
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblotrate 
      Caption         =   "otrate"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "rate per hour"
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "age"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   480
      TabIndex        =   11
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "empgender"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "empname"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Txtcmpempid 
      Caption         =   "cmpempid"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmbempid_Click()
rs.Open "select * from emptbl where empid=" & cmbempid.Text, con, 1, 2
Txtempname.Text = rs!empname
Txtempgender.Text = rs!empgender
Txtempage.Text = rs!empage
Txtrph.Text = rs!rph
Txtotrate.Text = rs!otrate
rs.Close
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\payroll.mdb;Persist Security Info=False "
con.Open
rs.Open "select empid from emptbl", con, 1, 2
While Not (rs.EOF)
cmbempid.AddItem (rs!empid)
rs.MoveNext
Wend
rs.Close
End Sub
Private Sub cmdcal_Click()
Txtgrosssal.Text = Val(Txtrph.Text) * Val(Txtwkhrs.Text) + Val(Txtotrate.Text) * Val(Txtothrs.Text)
Txttax.Text = Val(Txtgrosssal.Text) * 0.08
txtnetsal.Text = Val(Txtgrosssal.Text) - Val(Txttax.Text)
rs.Open "select * from emptbl where empid=" & cmbempid.Text, con, 1, 2
rs!netsal = Val(txtnetsal.Text)
rs.Update
rs.Close
End Sub

