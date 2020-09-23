VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Folders"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banned folders"
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   7815
      Begin VB.ListBox List2 
         Height          =   1635
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   960
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   6240
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "All folders listed here that are checked will not be searched. This includes sub-directories of those folders."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Folders"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CheckBox Check1 
         Caption         =   "Search Sub-directories of these folders"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3960
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.ListBox List1 
         Height          =   2085
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove"
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   $"Form6.frx":0000
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label Label3 
         Caption         =   $"Form6.frx":00B2
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   3240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function UpdateControls()
If List1.ListCount = 0 Then Command2.Enabled = False Else Command2.Enabled = True
If List2.ListCount = 0 Then Command6.Enabled = False Else Command6.Enabled = True
If List1.ListCount = 0 Then Command3.Enabled = False Else Command3.Enabled = True
If List2.ListCount = 0 Then Command7.Enabled = False Else Command7.Enabled = True
If List1.ListIndex <> -1 Then Command2.Enabled = True Else Command2.Enabled = False
If List2.ListIndex <> -1 Then Command6.Enabled = True Else Command6.Enabled = False
If Text1 = "" Then Command1.Enabled = False Else Command1.Enabled = True
If Text2 = "" Then Command5.Enabled = False Else Command5.Enabled = True

End Function
Private Sub Command1_Click()
If Dir(AddSlash(Text1)) = False Then
MsgBox "Path not found", vbCritical
Else
If List1.ListCount >= MaxFolders Then
MsgBox "Too many folders"
Else
For i = 0 To List1.ListCount - 1
If LCase(List1.List(i)) = LCase(AddSlash(Text1)) Then MsgBox "Folder already in list": Exit Sub
If Len(List1.List(i)) <= Len(AddSlash(Text1)) Then
If LCase(Left(AddSlash(Text1), Len(List1.List(i)))) = LCase(List1.List(i)) Then
MsgBox "The folder will already be searched through another folder"
Exit Sub
End If
End If
If Len(AddSlash(Text1)) <= Len(List1.List(i)) Then
If LCase(Left(List1.List(i), Len(AddSlash(Text1)))) = LCase(AddSlash(Text1)) Then
MsgBox "Adding this folder will cause '" + List1.List(i) + "' to be searched twice."
Exit Sub
End If
End If
Next

List1.AddItem AddSlash(Text1)
List1.Selected(List1.ListCount - 1) = True

End If
End If
UpdateControls
End Sub

Private Sub Command10_Click()
If Check1.Value = vbChecked Then SSD = True Else SSD = False
Dim i As Integer
SFC = List1.ListCount - 1
For i = 0 To List1.ListCount - 1
SearchFolder(i).FolderPath = List1.List(i)
SearchFolder(i).Used = List1.Selected(i)
Next


BFC = List2.ListCount - 1
For i = 0 To List2.ListCount - 1
BannedFolder(i).FolderPath = List2.List(i)
BannedFolder(i).Used = List2.Selected(i)

Next

List1.Clear
List2.Clear
UpdateControls
Me.Hide


End Sub

Private Sub Command2_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
If List1.ListCount <> 0 Then List1.ListIndex = 0
UpdateControls
End Sub

Private Sub Command3_Click()
List1.Clear
UpdateControls
End Sub

Private Sub Command4_Click()
Dim t As String
t = Browse(Me.hwnd, "Select folder")
If t <> "" Then Text1 = t: Command1_Click

UpdateControls
End Sub

Private Sub Command5_Click()
If Dir(AddSlash(Text2)) = False Then
MsgBox "Path not found", vbCritical
Else
If List2.ListCount >= MaxFolders Then
MsgBox "Too many folders"
Else
For i = 0 To List2.ListCount - 1
If LCase(List2.List(i)) = LCase(AddSlash(Text2)) Then MsgBox "Folder already in list": Exit Sub
If Len(List2.List(i)) <= Len(AddSlash(Text2)) Then
If LCase(Left(AddSlash(Text2), Len(List2.List(i)))) = LCase(List2.List(i)) Then
MsgBox "The folder will already be searched through another folder"
Exit Sub
End If
End If
'If Len(AddSlash(Text2)) <= Len(List2.List(i)) Then
'If LCase(Left(List2.List(i), Len(AddSlash(Text2)))) = LCase(AddSlash(Text2)) Then
'MsgBox "Adding this folder will cause '" + List2.List(i) + "' to be searched twice."
'Exit Sub
'End If
'End If
Next


List2.AddItem AddSlash(Text2)
List2.Selected(List2.ListCount - 1) = True

End If
End If
UpdateControls
End Sub

Private Sub Command6_Click()

List2.RemoveItem List2.ListIndex
If List2.ListCount <> 0 Then List2.ListIndex = 0
UpdateControls
End Sub

Private Sub Command7_Click()
List2.Clear
UpdateControls

End Sub

Private Sub Command8_Click()
Dim t As String
t = Browse(Me.hwnd, "Select destination folder")
If t <> "" Then Text2 = t: Command5_Click

UpdateControls
End Sub

Private Sub Command9_Click()

List1.Clear
List2.Clear
UpdateControls
Me.Hide
End Sub

Private Sub Form_Load()
SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, 1000, 0
SendMessage List2.hwnd, LB_SETHORIZONTALEXTENT, 1000, 0
Me.Icon = Form1.Icon
UpdateControls
End Sub

Private Sub List1_Click()
UpdateControls
End Sub

Private Sub Text1_Change()

UpdateControls
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command1_Click
End Sub

Private Sub Text2_Change()

UpdateControls
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Command5_Click


End Sub
