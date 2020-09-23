VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export List"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Filename only"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Export list to file"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Export list to clipboard"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   4335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BPressed = 1
If Option1.Value = True Then TypeOfCopy = 1 Else TypeOfCopy = 2
If Check1.Value = vbChecked Then FNOnly = True Else FNOnly = False
Me.Hide
End Sub

Private Sub Command2_Click()
BPressed = 2
Me.Hide
End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

BPressed = 2
End Sub
