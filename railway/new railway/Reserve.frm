VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form ve 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   13095
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      Height          =   1215
      Left            =   10680
      TabIndex        =   16
      Top             =   600
      Width           =   2655
   End
   Begin MSACAL.Calendar doj 
      Height          =   3495
      Left            =   10200
      TabIndex        =   15
      Top             =   6000
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   6165
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2014
      Month           =   3
      Day             =   10
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   7560
      Top             =   4680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\1180042\railway.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\1180042\railway.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.TextBox txtdoj 
      Height          =   735
      Left            =   3600
      TabIndex        =   14
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox trno 
      Height          =   975
      Left            =   3600
      TabIndex        =   13
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtpnrno 
      Height          =   855
      Left            =   10560
      TabIndex        =   12
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdbook 
      Caption         =   "BOOK TICKET"
      Height          =   1215
      Left            =   7200
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtdest 
      Height          =   975
      Left            =   3600
      TabIndex        =   10
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox txtsource 
      Height          =   975
      Left            =   3600
      TabIndex        =   9
      Top             =   5160
      Width           =   2655
   End
   Begin VB.TextBox txtage 
      Height          =   975
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtname 
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "PNR NUMBER"
      Height          =   735
      Left            =   7560
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "DESTINATION"
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "SOURCE"
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "TRAINNO"
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "DOJ"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "AGE"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "NAME"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "ve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

Private Sub cmdback_Click()
Unload Me
Form1.Show
End Sub

Private Sub cmdbook_Click()
Dim pnrno As String
Dim trainno As Integer
Dim seatno As Integer
Dim status As String
Dim capacity As Integer
Dim count As Integer
trainno = Val(trno.Text)
rs.Open " select capacity from traintbl where trainno = " & trainno, con, 1, 2
capacity = rs!capacity
rs.Close
rs1.Open "select * from reservetbl", con, 1, 2
count = rs1.RecordCount
If (count < capacity) Then
seatno = count + 1
status = "C"
Else
seatno = count - trcapacity + 1
status = "W"
End If
rs1.Close
rs1.Open "select * from reservetbl", con, 1, 2
rs1.AddNew
rs1!Name = txtname.Text
rs1!age = txtage.Text
rs1!doj = txtdoj.Text
rs1!Source = txtsource.Text
rs1!destination = txtdest.Text
pnrno = trainno & doj & seatno & status
txtpnrno = pnrno
rs1!ticketno = txtpnrno
rs1.Update
rs1.Close

End Sub

Private Sub doj_Click()
txtdoj.Text = doj.Value
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180022\database\railway.mdb;Persist Security Info=False"
con.Open
End Sub
