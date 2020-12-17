VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtcustid 
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "back to selection"
      Height          =   1335
      Left            =   10560
      TabIndex        =   8
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton cmdaddnew 
      Caption         =   "add new customer"
      Height          =   1455
      Left            =   6480
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Txtintdep 
      Height          =   735
      Left            =   9360
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Txtdop 
      Height          =   735
      Left            =   9480
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Txtconno 
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox Txtcustgender 
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox Txtcustage 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtcustname 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtaccno 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label intdeplbl 
      Caption         =   "initial deposit"
      Height          =   615
      Left            =   7560
      TabIndex        =   16
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label doplbl 
      Caption         =   "date of opening"
      Height          =   615
      Left            =   7560
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label contactnolbl 
      Caption         =   "contact no"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label custgenderlbl 
      Caption         =   "customer gender"
      Height          =   735
      Left            =   840
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label custagelbl 
      Caption         =   "customer age"
      Height          =   615
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label custidlbl 
      Caption         =   "customer id"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label custnamelbl 
      Caption         =   "customer name"
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label accnolbl 
      Caption         =   "account no"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
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
rs.Open "select * from custtbl", con, 1, 2
rs.movelast
Txtcustid.Text = (rs!customerid) + 1
Txtaccno.Text = (rs!accountno) + 1
rs.Close
Txtdop.Text = Format(Date, "dd-mm-yyyy")
End Sub

Private Sub cmdaddnew_Click()
rs.Open "select * from custtbl", con, 1, 2
rs.addnew
rs!customerid = Val(Txtcustid.Text)
rs!accountno = Val(Txtaccno.Text)
rs!customername = (txtcustname.Text)
rs!gender = (Txtcustgender.Text)
rs!contactno = Val(Txtconno.Text)
rs!dop = Val(Txtdop.Text)
rs!customerage = Val(Txtcustage.Text)
rs!balance = Val(Txtintdep.Text)
rs.Update
rs.Close

End Sub

Private Sub cmdexit_Click()
Unload Me
Form2.Show
End Sub



