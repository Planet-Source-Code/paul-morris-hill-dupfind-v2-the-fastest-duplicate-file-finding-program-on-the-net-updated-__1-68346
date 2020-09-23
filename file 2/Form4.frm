VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check6 
      Caption         =   "System files"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Read-Only files"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Hidden files"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Ask"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Zero byte files"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Crash protection"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4200
      TabIndex        =   6
      Text            =   "50"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Text            =   "*.*"
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Search system files"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "Search read-only files"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Search hidden files"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Ask before finding duplicate files"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Count zero byte files as duplicates"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   $"Form4.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   $"Form4.frx":0098
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "File Type"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"Form4.frx":0151
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Dim f As Integer
If Check1.Value = vbUnchecked Then
f = MsgBox("Warning: Removing crash protection can cause the program to become unstable or non responsive. Are you sure want to remove crash protection?", vbYesNo)
If f = vbYes Then Exit Sub Else Check1.Value = vbChecked
End If
End Sub

Private Sub Command1_Click()
Dim t
t = Val(Text1)
Text1.Text = Trim(Str(t))

If t > 900000 Then MsgBox "The maximum number of results allowed is 900000": Text1 = "900000": Exit Sub
If t <= 0 Then t = 900000
Maxr = t
Form1.File1.Pattern = Text2
t = Val(Text3)
If t < 1 Then MsgBox "Please choose a number between 1 and 900000": Text3 = "1": Exit Sub
If t > 900000 Then MsgBox "The maximum number of byte checks allowed is 900000": Text3 = "900000": Exit Sub
NOB = t
If Check1.Value = vbChecked Then CP = True Else CP = False
If Check2.Value = vbChecked Then ZBFAD = True Else ZBFAD = False
If Check3.Value = vbChecked Then AskPerm = True Else AskPerm = False
If Check4.Value = vbChecked Then Form1.File1.Hidden = True Else Form1.File1.Hidden = False
If Check5.Value = vbChecked Then Form1.File1.ReadOnly = True Else Form1.File1.ReadOnly = False
If Check6.Value = vbChecked Then Form1.File1.System = True Else Form1.File1.System = False

Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon

End Sub
