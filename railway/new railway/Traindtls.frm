VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   15705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdbook 
      Caption         =   "book"
      Height          =   1215
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   600
      Top             =   6720
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2143
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180022\database\railway.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180022\database\railway.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "traintbl"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "SHOW"
      Height          =   1455
      Left            =   4680
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   5520
      Left            =   6840
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   600
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdbook_Click()
Unload Me
ve.Show
End Sub

Private Sub cmdshow_Click()
Dim trname As String
trname = List1.Text
rs.Open "select * from Traintbl where Trainname ='" & trname & "'", con, 1, 2
While Not (rs.EOF)
MsgBox (rs.RecordCount)
List2.AddItem ("*****************************")
List2.AddItem ("Tarin Name :" & rs!Trainname)
List2.AddItem ("Tarin No :" & rs!trainno)
List2.AddItem ("Source  :" & rs!Source)
List2.AddItem ("Destination  :" & rs!Destn)
List2.AddItem ("*****************************")
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180022\database\railway.mdb;Persist Security Info=False"
con.Open
rs.Open "select distinct Trainname from Traintbl", con, 1, 2
While Not (rs.EOF)
List1.AddItem (rs!Trainname)
rs.MoveNext
Wend
rs.Close
End Sub

