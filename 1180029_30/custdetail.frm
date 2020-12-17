VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdexit 
      Caption         =   "back to selection"
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   4800
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   8
      FormatString    =   "customerid|accountno|customername|gender|customerage|contactno|dop|balance"
   End
   Begin VB.ComboBox cmbcustid 
      Height          =   315
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label custlbl 
      Caption         =   "customer id"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
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
rs.Open " select customerid from custtbl", con, 1, 2
While Not (rs.EOF)
cmbcustid.AddItem (rs!customerid)
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub cmbcustid_Click()
rs.Open "select * from custtbl where customerid=" & cmbcustid.Text, con, 1, 2
MSFlexGrid1.AddItem (rs!accountno & vbTab & rs!customername & vbTab & rs!gender & vbTab & rs!customerage & vbTab & rs!contactno & vbTab & rs!dop & vbTab & rs!balance)
rs.Close
End Sub
Private Sub cmdexit_Click()
Unload Me
Form2.Show
End Sub

