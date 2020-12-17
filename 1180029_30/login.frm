VERSION 5.00
Begin VB.Form form4 
   Caption         =   "login"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   6360
      TabIndex        =   5
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "login"
      Height          =   855
      Left            =   3720
      TabIndex        =   4
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtupaswd 
      Height          =   1095
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtuname 
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblupaswd 
      Caption         =   "password"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lbluname 
      Caption         =   "username"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Dim rs As ADODB.Recordset
Private Sub Form_Load()
Set con = New ADODB.Connection
con.ConnectionString = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\student.mdb;Persist Security Info=False"
con.Open
End Sub

Private Sub cmdlogin_Click()
Set rs = New ADODB.Recordset
rs.Open "select * from logintbl", con, 1, 2
While Not (rs.EOF)
If (txtuname.Text = rs!uname) And (txtupaswd.Text = rs!upaswd) Then
MsgBox ("login successful")
rs.Close
Unload Me
form1.Show
Exit Sub
Else
rs.MoveNext
End If
Wend
rs.Close
MsgBox ("login failed")
End Sub
Private Sub cmdcancel_Click()
txtuname.Text = " "
txtupaswd.Text = " "
End Sub


