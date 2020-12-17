VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form form3 
   Caption         =   "showselect"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2895
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   7
      FormatString    =   $"showselect.frx":0000
   End
   Begin VB.CommandButton Cmdgo 
      Caption         =   "go"
      Height          =   1575
      Left            =   12960
      TabIndex        =   2
      Top             =   3960
      Width           =   4455
   End
   Begin VB.ComboBox cmbdept 
      Height          =   315
      Left            =   11640
      TabIndex        =   1
      Text            =   "nill"
      Top             =   840
      Width           =   6975
   End
   Begin VB.ComboBox cmbyear 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "00"
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "department"
      Height          =   615
      Left            =   9960
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "year"
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection

Dim rs As ADODB.Recordset

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\student.mdb;Persist Security Info=False"
con.Open
rs.Open "select distinct studdept from studtbl", con, 1, 2
While Not (rs.EOF)
cmbdept.AddItem (rs!studdept)
rs.MoveNext
Wend
rs.Close
rs.Open "select distinct studyear from studtbl", con, 1, 2
While Not (rs.EOF)
cmbyear.AddItem (rs!studyear)
rs.MoveNext
Wend
rs.Close
End Sub


Private Sub Cmdgo_Click()
MSFlexGrid1.Rows = 1
If (cmbdept.Text = "nill" And Val(cmbyear.Text) = 0) Then
MsgBox ("please select any one option")
ElseIf (cmbdept.Text = "nill" And Val(cmbyear.Text) <> 0) Then
rs.Open " select * from studtbl where studyear=" & cmbyear.Text, con, 1, 2
ElseIf (cmbdept.Text <> "nill" And Val(cmbyear.Text) = 0) Then
rs.Open " select * from studtbl where studdept='" & cmbdept.Text & "'", con, 1, 2
ElseIf (cmbdept.Text <> "nill" And Val(cmbyear.Text) <> 0) Then
rs.Open " select * from studtbl where studyear= " & cmbyear.Text & " and studdept='" & cmbdept.Text & "'", con, 1, 2
End If
While Not (rs.EOF)
MSFlexGrid1.AddItem (rs!studrollno & vbTab & rs!studname & vbTab & rs!studgender & vbTab & rs!studage & vbTab & rs!studcontactno & vbTab & rs!studdept & vbTab & rs!studyear)
rs.MoveNext
Wend
rs.Close
End Sub

