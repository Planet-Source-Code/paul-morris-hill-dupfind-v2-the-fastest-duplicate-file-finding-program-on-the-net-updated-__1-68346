VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "View List"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   1320
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Menu mnuclick 
      Caption         =   "clickmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuopncont 
         Caption         =   "Open containing folder"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete file"
      End
      Begin VB.Menu mnucop 
         Caption         =   "Copy File"
      End
      Begin VB.Menu mnumovfile 
         Caption         =   "MoveFile"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function FileOperation(Operation As Integer)
On Error GoTo errtrap
Dim ll As String
Dim temp As String
Dim ext As String
If List1.ListIndex = -1 Then Exit Function
ll = List1.List(List1.ListIndex) ' list list
If Dir(ll) <> True Then Exit Function
If Operation = 1 Then
    Shell "explorer.exe """ + GetPath(ll) + """", vbNormalFocus
    Exit Function
    'open containing folder
End If
If Operation = 4 Then
    If MsgBox("Are you sure you want to delete this file", vbYesNo) = vbYes Then
        Kill ll
        MsgBox "File deleted"
    End If
    Exit Function
End If
GetFileEXT GetFileName(ll), temp, ext
cd.Filter = "*" + ext + " |" + ext
cd.ShowSave
If cd.Filename = "" Then Exit Function
If Dir(cd.Filename) = True Then MsgBox "File already exists": Exit Function
Select Case Operation
Case 2
FileCopy ll, cd.Filename
MsgBox "File copied"
Case 3
Name ll As cd.Filename
MsgBox "File moved"
End Select
Exit Function
errtrap:
MsgBox "Error: " + Str(Err.Number) + " " + Err.Description, vbCritical
Exit Function
End Function

Private Sub Form_Resize()
On Error Resume Next
List1.Width = Me.Width - 125
List1.Height = Me.Height - 400
Form2.Icon = Form1.Icon
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnuclick
End Sub

Private Sub mnucop_Click()
FileOperation 2

End Sub

Private Sub mnudelete_Click()
FileOperation 4
End Sub

Private Sub mnumovfile_Click()
FileOperation 3
End Sub

Private Sub mnuopncont_Click()
FileOperation 1
End Sub
