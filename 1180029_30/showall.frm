VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form form2 
   Caption         =   "showall"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdback 
      Caption         =   "back to selection"
      Height          =   1215
      Left            =   4800
      TabIndex        =   1
      Top             =   7080
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "showall.frx":0000
      Height          =   2895
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   8
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0)._NumMapCols=   7
      _Band(0)._MapCol(0)._Name=   "studrollno"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "studname"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "studage"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
      _Band(0)._MapCol(3)._Name=   "studyear"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Alignment=   7
      _Band(0)._MapCol(4)._Name=   "studgender"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "studcontactno"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(5)._Alignment=   7
      _Band(0)._MapCol(6)._Name=   "studdept"
      _Band(0)._MapCol(6)._RSIndex=   6
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1815
      Left            =   3600
      Top             =   4800
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3201
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\1180029&30\student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "studtbl"
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
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdback_Click()
Unload Me
form1.Show
End Sub
