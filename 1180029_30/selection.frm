VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   1095
      Left            =   8760
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdtransact 
      Caption         =   "transaction"
      Height          =   1335
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdaddnew 
      Caption         =   "add new customer"
      Height          =   1335
      Left            =   8520
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdcustdtls 
      Caption         =   "customer details"
      Height          =   1455
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection

Dim rs As ADODB.Recordset


Private Sub cmdaddnew_Click()
Unload Me
Form4.Show
End Sub

Private Sub cmdcustdtls_Click()
Unload Me
Form3.Show
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdtransact_Click()
Unload Me
Form5.Show
End Sub
