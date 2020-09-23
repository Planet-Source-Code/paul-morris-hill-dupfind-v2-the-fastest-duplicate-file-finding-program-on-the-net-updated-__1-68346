Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const LB_SETHORIZONTALEXTENT As Long = &H194
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public SSD              As Boolean ' search sub-directories
Public Maxr             As Long ' max results
Public CP               As Boolean ' crash protection
Public NOB              As Long ' No of Bytes
Public BPressed         As Integer
Public TypeOfCopy       As Integer
Public FNOnly           As Boolean
Public ZBFAD            As Boolean ' zero byte files are duplicates
Public CurrentType      As Integer
Public Const MaxFolders As Integer = 100
Public NoOfTimes        As Integer
Public SFC              As Long ' search folder count
Public BFC              As Long ' Banned folder count
Public AskPerm          As Boolean
Type BannedFolders
FolderPath              As String
Used                    As Boolean
Count                   As Long
End Type
Type SearchFolders
FolderPath  As String
Used        As Boolean
Count       As Long
End Type
Public SearchFolder(101) As SearchFolders
Public BannedFolder(101) As BannedFolders
Public Function Browse(mhwnd As Long, message As String) As String



Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

szTitle = message

With tBrowseInfo
   .hWndOwner = mhwnd ' Owner Form
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   Browse = sBuffer
End If


End Function


Public Function Dir(strFile As String) As Boolean


    If PathFileExists(strFile) = 1 Then
        Dir = True
    ElseIf PathFileExists(strFile) = 0 Then
        Dir = False
    End If
End Function

Public Function AddSlash(Path As String)
If Right(Path, 1) <> "\" And Right(Path, 1) <> "/" Then
    AddSlash = Path + "\"
Else
    AddSlash = Path
End If
End Function

Public Function GetFileEXT(mFileName As String, Begin As String, Extension As String)
On Error Resume Next
Dim j    As Long
Dim mlen As Long

For j = Len(mFileName) To 1 Step -1
    If Mid(mFileName, j, 1) = "." Then mlen = j: Exit For
Next

Begin = Left(mFileName, mlen - 1)
Extension = Right(mFileName, Len(mFileName) - mlen + 1)
End Function

Public Function GetFileName(Filename As String) As String
Dim h As Long

For h = Len(Filename) To 1 Step -1
    If Mid(Filename, h, 1) = "\" Then
        GetFileName = Right(Filename, Len(Filename) - h): Exit For
    End If
Next

End Function

Public Function GetPath(Filename As String) As String
Dim h As Long

For h = Len(Filename) To 1 Step -1
    If Mid(Filename, h, 1) = "\" Then
        GetPath = Left(Filename, h): Exit For
    End If
Next
End Function
