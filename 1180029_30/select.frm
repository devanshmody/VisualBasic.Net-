VERSION 5.00
Begin VB.Form form1 
   Caption         =   "select"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "form1"
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgo 
      Caption         =   "go"
      Height          =   1215
      Left            =   4440
      TabIndex        =   2
      Top             =   5520
      Width           =   5535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "show selected students"
      Height          =   1215
      Left            =   4320
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "show all students"
      Height          =   1215
      Left            =   4320
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset



Private Sub Cmdgo_Click()
If (Check1.Value = 1 And Check2.Value = 0) Then
Unload Me
form2.Show
End If

If (Check2.Value = 1 And Check1.Value = 0) Then
Unload Me
form3.Show
End If

If (Check2.Value = 1 And Check1.Value = 1) Then
MsgBox ("please select one item only")
Check2.Value = 0
Check1.Value = 0
End If

If (Check1.Value = 0 And Check2.Value = 0) Then
MsgBox ("please select atleast one option")
End If


End Sub


