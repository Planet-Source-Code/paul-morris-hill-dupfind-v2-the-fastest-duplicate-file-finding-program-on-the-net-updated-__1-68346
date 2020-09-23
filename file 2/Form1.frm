VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DupFind"
   ClientHeight    =   3540
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   6600
      TabIndex        =   19
      Top             =   1320
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Export List"
      Filter          =   "Text Files *.txt | *.txt"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8235
      TabIndex        =   7
      Top             =   500
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pr1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   8235
      TabIndex        =   4
      Top             =   75
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   3720
      ReadOnly        =   0   'False
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label12 
      Caption         =   "0 files per minute"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Rate:"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "0s"
      Height          =   255
      Left            =   4320
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Elapsed time:"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Folders Searched"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Files found so far:"
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Searched folders"
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Files being searched"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Current folder being searched"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   9120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnucopnon 
         Caption         =   "Copy Non-Duplicates"
      End
      Begin VB.Menu mnucopdup 
         Caption         =   "Copy Duplicates"
      End
      Begin VB.Menu mnuOOE 
         Caption         =   "Copy One of each duplicate and all non-duplicates"
      End
      Begin VB.Menu mnumovnon 
         Caption         =   "Move Non-Duplicates"
      End
      Begin VB.Menu movdup 
         Caption         =   "Move Duplicates"
      End
      Begin VB.Menu mnud4 
         Caption         =   "-"
      End
      Begin VB.Menu mnudeldup 
         Caption         =   "Delete Duplicates"
      End
      Begin VB.Menu mnudelnon 
         Caption         =   "Delete Non-Duplicates"
      End
      Begin VB.Menu mnud2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools"
      Begin VB.Menu mnuerr 
         Caption         =   "Export error log"
      End
      Begin VB.Menu mnuexpff 
         Caption         =   "Export files found list"
      End
      Begin VB.Menu mnuexpdf 
         Caption         =   "Export duplicate file list"
      End
      Begin VB.Menu mnunond 
         Caption         =   "Export non-duplicate file list"
      End
      Begin VB.Menu mnud3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuviewnon 
         Caption         =   "View non-duplicate file list"
      End
      Begin VB.Menu mnuviewdup 
         Caption         =   "View duplicate file list"
      End
      Begin VB.Menu dsh2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusss 
         Caption         =   "Show Search Summary"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnud 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopt 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "Search"
      Begin VB.Menu mnustrt 
         Caption         =   "Start Search"
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop Search"
      End
      Begin VB.Menu mnupause 
         Caption         =   "Pause"
      End
      Begin VB.Menu dsh 
         Caption         =   "-"
      End
      Begin VB.Menu much 
         Caption         =   "Choose search folders"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Type mSearchResult
Filename    As String
FileSize    As Single
IsDuplicate As Boolean

End Type

Private Type mSearchResult2
Filename        As String
FileSize        As Single
DuplicateNum    As Long
End Type
Private Type mCopyList
mFileName As String
End Type

Dim ErrorCount          As Long
Dim ErrorMessage        As String
Dim PauseSearch         As Boolean
Dim mCopyFile(900001)   As mCopyList
Dim ttc                 As Long
Dim mSearchCount        As Long
Dim mSearchRs(900001)   As mSearchResult
Dim mDuplicate(900001)  As mSearchResult2
Dim mDupCount           As Long
Dim mDupGroupCount      As Long
Dim StopSearch          As Boolean

Private Function NonDupAction(Action As Integer)
On Error GoTo errtrap
Dim PathToCopy As String
Dim i, ind, CopyErrorCount As Long
Dim CopyFile As String
Dim CopyWholeFileName As String
Dim CopyFileEXT As String
Dim CopyFileName As String
Dim EndFileName As String
Dim strIND
If Action = 3 Then
Form2.Show

Form2.List1.Clear
PathToCopy = "Somefakevalue"
Else
PathToCopy = Browse(Me.hwnd, "Select destination folder")
End If
If PathToCopy = "" Then Exit Function
pr1.Max = 100
PathToCopy = AddSlash(PathToCopy)

For i = 0 To mSearchCount
pr1.Value = (100 / mSearchCount) * i
If mSearchRs(i).IsDuplicate = True Then GoTo mNext
CopyFile = mSearchRs(i).Filename
    CopyWholeFileName = GetFileName(CopyFile)
         GetFileEXT CopyWholeFileName, CopyFileName, CopyFileEXT

If CopyFile = "" Then GoTo mNext
Retry:
If ind <> 0 Then strIND = " " + "(" + Str(ind) + ")" Else strIND = ""
EndFileName = PathToCopy + CopyFileName + strIND + CopyFileEXT
If Dir(EndFileName) Then ind = ind + 1: GoTo Retry Else ind = 0
Select Case Action
Case 1
FileCopy CopyFile, EndFileName
Case 2
Name CopyFile As EndFileName
Case 3
Form2.List1.AddItem CopyFile
End Select
mNext:
'temp = GetTickCount()

If CP = True Then DoEvents

'aft = GetTickCount()
'd = d + (aft - temp)
Next
Dim errstring As String
Dim errmes As String
errstring = "Errors: %"
MsgBox Replace(errstring, "%", Str(CopyErrorCount)) + errmes
pr1.Value = 0
Exit Function
errtrap:

CopyErrorCount = CopyErrorCount + 1
errmes = errmes + vbCrLf + Err.Description
Err.Clear
Resume Next

End Function

Private Function DuplicateAction(Action As Integer)
On Error GoTo errtrap
Dim PathToCopy          As String
Dim i                   As Long
Dim ind                 As Long
Dim dupnum              As Long
Dim CopyErrorCount      As Long
Dim CopyFile            As String
Dim CopyWholeFileName   As String
Dim CopyFileEXT         As String
Dim CopyFileName        As String
Dim EndFileName         As String
Dim strIND              As String
Dim OldDup              As Integer
OldDup = -1
If Action = 3 Then
Form2.List1.Clear
Form2.Show

PathToCopy = "Somefakevalue"
Else
PathToCopy = Browse(Me.hwnd, "Select destination folder")
End If
If PathToCopy = "" Then Exit Function
PathToCopy = AddSlash(PathToCopy)
pr1.Max = 100
For i = 0 To mDupCount
    pr1.Value = (100 / mDupCount) * i

    CopyFile = mDuplicate(i).Filename
    CopyWholeFileName = GetFileName(CopyFile)
    GetFileEXT CopyWholeFileName, CopyFileName, CopyFileEXT
    dupnum = mDuplicate(i).DuplicateNum
    If CopyFile = "" Then GoTo mNext
Retry:
    If ind <> 0 Then strIND = " " + "(" + Str(ind) + ")" Else strIND = ""
    EndFileName = PathToCopy + Str(dupnum) + " " + CopyFileName + strIND + CopyFileEXT
    If Dir(EndFileName) Then ind = ind + 1: GoTo Retry Else ind = 0
    
    Select Case Action
    Case 1
    FileCopy CopyFile, EndFileName
    Case 2
    Name CopyFile As EndFileName
    Case 3
    Load Form2
    If dupnum <> OldDup Then
    Form2.List1.AddItem "Group " + Trim(Str(dupnum)) + " File size: " + Str(mDuplicate(i).FileSize)
    OldDup = dupnum
    End If
    Form2.List1.AddItem (CopyFile)
    End Select
mNext:
Next
Dim errstring As String
Dim errmes As String
errstring = "Errors: %" ' Unneccesary
MsgBox Replace(errstring, "%", Str(CopyErrorCount)) + errmes
pr1.Value = 0
Exit Function
errtrap:

CopyErrorCount = CopyErrorCount + 1
errmes = errmes + vbCrLf + Err.Description
Err.Clear
Resume Next
End Function
Private Function Swap(first As Long, Second As Long)
Dim bufn As String
Dim bufs As Single

bufs = mSearchRs(first).FileSize
bufn = mSearchRs(first).Filename
mSearchRs(first).Filename = mSearchRs(Second).Filename
mSearchRs(first).FileSize = mSearchRs(Second).FileSize
mSearchRs(Second).Filename = bufn
mSearchRs(Second).FileSize = bufs
End Function
Private Function GetTimeFromSeconds(Seconds As Long) As String
Dim Minutes As Integer
Dim Hours   As Integer
If Seconds > 59 Then
    Minutes = Int(Seconds \ 60)
    Seconds = Seconds - Int(Seconds \ 60) * 60
Else
    GetTimeFromSeconds = Trim(Str(Seconds)) + "s"
    Exit Function
End If
If Minutes > 59 Then
    Hours = Int(Minutes / 60)
    Minutes = Minutes - Int(Minutes / 60) * 60
Else
    GetTimeFromSeconds = Trim(Str(Minutes)) + "m " + Trim(Str(Seconds)) + "s"
    Exit Function
End If
GetTimeFromSeconds = Trim(Str(Hours)) + "h " + Trim(Str(Minutes)) + "m " + Trim(Str(Seconds)) + "s"


End Function




Private Sub FindDup(ResultCount As Long, d As Long)
On Error GoTo errortrap
Dim sof As Long
    sof = GetTickCount()
Dim temp      As Long
Dim dupfound  As Boolean
Dim aft       As Long
Dim mrate     As Long
Dim cnt       As Long
Dim location1 As Long
Dim Crnt      As Long
Dim i         As Long
Dim f         As Long
Dim r         As Long
Dim j         As Long
Dim u         As Long
Dim bufs      As Single
Dim Byt1      As Byte
Dim Byt2      As Byte
Dim fn        As Long
Dim fn2       As Long
Dim IsDup     As Boolean
Dim dupnum    As Long
Dim pls       As Long
Dim fcf       As Long
Dim ifs       As Single
Dim ipp       As Single
Dim ifn       As String
Dim ipf       As String

ErrorMessage = ""
Crnt = -1
dupnum = -1
ErrorCount = -1
mDupGroupCount = -1
pr1.Max = 100
Label1.Caption = "Arranging Files"
' Start of sort Start of sort Start of sort Start of sort Start of sort Start of sort Start of sort Start of sort Start of sort
Sort 0, mSearchCount, d
' end of sort
i = -1
pr1.Max = 100
sof = GetTickCount()
Dim exin As Long
Label1.Caption = "Finding duplicate files"
Do

    cnt = GetTickCount()
    If cnt - sof <> 0 Then
   ' calculate the file rate
        mrate = Int((i + pls) / (((cnt - sof) / 1000) / 60)) ' here again --------------------------
        'If avgcon > 100 Then
        'avgtot = avgtot + mrate
        'avgcon = avgcon + 1
        'avgr = avgtot / avgcon
        'mrate = avgr
        'End If
        'If mrate <> 0 Then
        'exin = (mSearchCount / (mrate / 60)) - r
        'If Int(exin) < 0 Then exin = 0


        'Label1.Caption = "Finding duplicate files: Estimated completion time: " _
          + GetTimeFromSeconds(mSearchCount / (mrate / 60)) + _
           " Estimated time left: " + GetTimeFromSeconds(exin)
        'End If
    End If
    Label12.Caption = Str(mrate) + " Files per min"
    If PauseSearch = True Then
        temp = GetTickCount()
        MsgBox "Press Ok to continue the search"
        PauseSearch = False
        aft = GetTickCount()
        d = d + (aft - temp)
    End If
    f = GetTickCount()
    r = (f - d) / 1000
    Label10.Caption = GetTimeFromSeconds(r)
    i = i + 1
    If i > ResultCount Then Exit Do
        pls = 1
        pr1.Value = (100 / (ResultCount + 1)) * i
'next search result
NSR:

    ifs = mSearchRs(i).FileSize
    ipp = mSearchRs(i + pls).FileSize
    ifn = mSearchRs(i).Filename
    ipf = mSearchRs(i + pls).Filename
    ' if the files are zero size and the zero byte option is on
    ' automaticly set them as duplicates
    If ifs = 0 And ipp = 0 And (i + pls) <= ResultCount And ipf <> ifn Then IsDup = ZBFAD: GoTo ZeroSize
    ' if they are teh same size goto byte checking
    If ifs = ipp And (i + pls) <= ResultCount And ipf <> ifn Then
        CurrentType = 1 ' starting pattern byte checking
        'through the number of bytes selected in teh options
ByteComparison:
        If CurrentType = 1 Then
        ' if the file size is smaller than the number of
        ' bytes then just search the whole file
            If mSearchRs(i).FileSize < NOB Then
                NoOfTimes = mSearchRs(i).FileSize
                CurrentType = 2
            Else
                NoOfTimes = NOB
            End If
        End If
        ' set the values so that the error log knows
        ' which file the error happened on
        fcf = i
        fn = FreeFile
        If Dir(mSearchRs(i).Filename) = False Then GoTo ZeroSize
        If Dir(mSearchRs(i + pls).Filename) = False Then GoTo ZeroSize
        Open mSearchRs(i).Filename For Binary As fn
        fcf = i + pls
        fn2 = FreeFile
        Open mSearchRs(i + pls).Filename For Binary As fn2
        'compare the bytes
        For j = 1 To NoOfTimes
            IsDup = True
  
            If CP = True Then DoEvents
            
            If mSearchRs(i).FileSize <> 0 Then
                If CurrentType = 1 Then
                'calculate the location in the file to check
                    location1 = Int((mSearchRs(i).FileSize / NoOfTimes) * j)
                    If location1 = 0 Then location1 = 1
                    'check the location wasnt put to
                    ' high because of the int function
                    If location1 > mSearchRs(i).FileSize Then location1 = mSearchRs(i).FileSize
                Else
                    location1 = j
                End If
   
            End If
    
            fcf = i
            Get #fn, location1, Byt1
            fcf = i + pls
            Get #fn2, location1, Byt2
            ' compare the two bytes
            If Byt2 <> Byt1 Then IsDup = False: Exit For
                 
        Next
        Close #fn
        Close #fn2
    Else
        IsDup = False
    End If
   ' if the file is zero size the program will skip to here
ZeroSize:
    If IsDup = True Then
        If CurrentType = 1 Then CurrentType = 2: GoTo ByteComparison
        dupfound = True
        'set the index for the next file to be checked
        pls = pls + 1
        ' look through the files again
        GoTo NSR
    ' if it wasnt a duplicate
    Else
    ' but if there was a duplicate group
        If dupfound = True Then
            dupfound = False
            'increment the duplicate group num
            dupnum = dupnum + 1
            'add them to the duplicate file array
            For u = i To (i + pls - 1) '
                Crnt = Crnt + 1
                mDuplicate(Crnt).Filename = mSearchRs(u).Filename
                mDuplicate(Crnt).FileSize = mSearchRs(u).FileSize
                mDuplicate(Crnt).DuplicateNum = dupnum
                mSearchRs(u).IsDuplicate = True
            Next
            i = i + (pls - 1)
        Else
            mSearchRs(u).IsDuplicate = False
        End If
    End If
    If CP = True Then DoEvents

Loop
mDupCount = Crnt
mDupGroupCount = dupnum
Label1.Caption = ""
Exit Sub
errortrap:
ErrorCount = ErrorCount + 1
ErrorMessage = ErrorMessage + vbCrLf + "Error " + Trim(Str(Err.Number)) + vbCrLf + Err.Description + vbCrLf + "File: " + mSearchRs(fcf).Filename

Resume Next
End Sub

Private Sub Command1_Click()
File1.Refresh
'check for search folders
If SFC = -1 Then MsgBox "Please choose at least 1 directory to search": Exit Sub
On Error GoTo errtrap
Dim d As Long
'reset progress bar
pr1.Value = 0
'set timer
d = GetTickCount()
'clear search folder list
List1.Clear
' declare variables
Dim DLI     As String
Dim cbf     As String
Dim g       As String
Dim temprs  As String
Dim temp    As Single
Dim aft     As Single
Dim FCount  As Long
Dim Result  As Long
Dim ret     As Integer
Dim A       As Integer
Dim c       As Long
Dim i       As Long
Dim z       As Long
Dim f       As Long
Dim r       As Long
Dim o       As Integer
Dim l1      As String
'set the result count to -1
Result = -1
'add all search folders to the list
For i = 0 To SFC
    If SearchFolder(i).Used = True Then
        List1.AddItem SearchFolder(i).FolderPath
    End If
Next
If List1.ListCount = 0 Then MsgBox "Please choose at least 1 directory to search": Exit Sub
' set path
Dir1.Path = List1.List(0)
'disable and enable controls
mnuFile.Enabled = False
mnusss.Enabled = False
mnupause.Enabled = True
Text1.Enabled = False  'Search Keywords
  'Bit Checks
Drive1.Enabled = False
Dir1.Enabled = False
File1.Enabled = False
List1.Enabled = False
mnuopt.Enabled = False
Command2.Enabled = True
Command1.Enabled = False 'Search Button
mnustrt.Enabled = False
mnustop.Enabled = True
' search for files loop

Do
    If PauseSearch = True Then
        temp = GetTickCount() ' pause timer
        MsgBox "Press Ok to continue the search"
        PauseSearch = False
        aft = GetTickCount()
        d = d + (aft - temp)
    End If
    If StopSearch = True Then

        temp = GetTickCount()
        A = MsgBox("Do you want to stop the search? " + Str(Result + 1) + " Results Found", vbYesNo)
        aft = GetTickCount()
        d = d + (aft - temp)
        If A = vbNo Then
            StopSearch = False
        Else
            StopSearch = False: Exit Do
        End If
    End If
    Dir1.Path = List1.List(z)
    FCount = FCount + 1
    Label8.Caption = Trim(Str(FCount))
    ' add the files in that directory to the files found array
    For c = 0 To File1.ListCount - 1
        f = GetTickCount()
        r = (f - d) / 1000
        Label10.Caption = GetTimeFromSeconds(r)
        If InStr(1, File1.List(c), Text1) <> 0 Then
            Result = Result + 1
            Label6.Caption = Trim(Str(Result + 1))
            g = AddSlash(Dir1.Path)
            mSearchRs(Result).Filename = g + File1.List(c)
            mSearchRs(Result).FileSize = FileLen(g + File1.List(c))
            If Result + 1 >= Maxr And Maxr <> 0 Then Exit Do
        End If
    Next
    If SSD = True Then ' S-earch S-ub D-irectories
        For i = 0 To Dir1.ListCount - 1
            DLI = LCase(AddSlash(Dir1.List(i)))
            ' banned folder check -----------------------------------------------
            For o = 0 To BFC
                cbf = LCase(BannedFolder(o).FolderPath)
                l1 = Left(DLI, Len(cbf))
                If Len(DLI) >= Len(cbf) And BannedFolder(o).Used = True Then
                    If l1 = cbf Then
                        GoTo IsBanned
                    End If
                End If
            Next

            List1.AddItem DLI
IsBanned:
        Next
    End If
    z = z + 1
    If z = List1.ListCount Then Exit Do
    If CP = True Then DoEvents ' crash protection
Loop
mSearchCount = Result
Command2.Enabled = False
mnustop.Enabled = False
' Find Duplicates Here
If AskPerm = True Then
    temprs = Trim(Str(mSearchCount + 1)) + " Results found, "
    temp = GetTickCount()
    ret = MsgBox("Do you want to find duplicates?", vbYesNo)
    aft = GetTickCount()
    d = d + (aft - temp)
Else
    ret = vbYes
End If
If ret = vbYes Then
    FindDup mSearchCount, d
Else
    mDupCount = -1
    mDupGroupCount = -1
    ErrorCount = -1
    ErrorMessage = ""
End If
f = GetTickCount()
r = (f - d) / 1000
Text1.Enabled = True  'Search Keywords
 
Drive1.Enabled = True
Dir1.Enabled = True
File1.Enabled = True
List1.Enabled = True
Command1.Enabled = True 'Search Button
mnustrt.Enabled = True

ttc = r ' time to complete
MsgBox "Search Complete " + vbCrLf + "Summary: " + vbCrLf + _
 Str(Result + 1) + " Matches " + vbCrLf + Str(mDupCount + 1) _
 + " Duplicates Found In " + GetTimeFromSeconds(r) + vbCrLf + Str(mDupGroupCount + 1) + " Duplicate Groups" + vbCrLf + Str(ErrorCount + 1) + " Errors "
mnuopt.Enabled = True
mnuFile.Enabled = True
mnusss.Enabled = True
mnupause.Enabled = False
pr1.Value = 0
Exit Sub
errtrap:
ErrorMessage = ErrorMessage + "Error: " + Str(Err.Number) + " " + Err.Description
ErrorCount = ErrorCount + 1
Resume Next
End Sub







Private Sub Command2_Click()

StopSearch = True
End Sub







Private Sub Dir1_Change()
On Error GoTo errtrap
File1.Path = Dir1.Path
Exit Sub
errtrap:
MsgBox "Error: " + Str(Err.Number) + " " + Err.Description, vbCritical

End Sub

Private Sub Drive1_Change()
On Error GoTo errtrap
Dir1.Path = Drive1.List(Drive1.ListIndex)
Exit Sub
errtrap:
MsgBox "Error: " + Str(Err.Number) + " " + Err.Description, vbCritical

End Sub

Private Sub Form_Load()
On Error Resume Next
SFC = -1
BFC = -1
SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, 1000, 0


ZBFAD = True
NOB = 50
CP = True
Maxr = 50000

mSearchCount = -1
File1.Pattern = "*.*"
Form1.Show
Form3.Show 1
Form6.Show 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
End
End Sub



Private Sub mnucopdup_Click()
DuplicateAction 1
End Sub

Private Sub mnucopnon_Click()
NonDupAction 1
End Sub
Private Sub mnudeldup_Click()
On Error GoTo errtrap
Dim i As Long
Dim A As String
Dim errCount As Long
Dim errmes As String
A = InputBox("Are you sure you want to delete all duplicates? Type 'delete' to continue")
If A <> "delete" Then MsgBox "No files deleted": Exit Sub
pr1.Max = 100

For i = 0 To mDupCount
pr1.Value = (100 / mDupCount) * i
Kill mDuplicate(i).Filename
Next
MsgBox "Errors: " + Str(errCount) + errmes
pr1.Value = 0
Exit Sub
errtrap:
errCount = errCount + 1
errmes = errmes + vbCrLf + Err.Description
Resume Next
End Sub

Private Sub mnudelnon_Click()
On Error GoTo errtrap
Dim i As Long
Dim A As String
Dim errCount As Long
Dim errmes As String
A = InputBox("Are you sure you want to delete all non-duplicates? Type 'delete' to continue")
If A <> "delete" Then MsgBox "No files deleted": Exit Sub
pr1.Max = 100

For i = 0 To mSearchCount
pr1.Value = (100 / mSearchCount) * i
If mSearchRs(i).IsDuplicate = False Then
Kill mSearchRs(i).Filename
End If
Next
MsgBox "Errors: " + Str(errCount) + errmes
pr1.Value = 0
Exit Sub
errtrap:
errCount = errCount + 1
errmes = errmes + vbCrLf + Err.Description
Resume Next
End Sub

Private Sub mnuerr_Click()
Form5.Check1.Enabled = False
Form5.Check1.Value = vbUnchecked
Form5.Show 1
Form5.Check1.Enabled = True
Dim clist As String
Dim fnum As Long
clist = ErrorMessage
If BPressed = 2 Then Exit Sub
 If TypeOfCopy = 1 Then Clipboard.Clear: Clipboard.SetText clist: MsgBox "Error log copied to clipboard succesfully"
    ' if it is to export to file
    If TypeOfCopy = 2 Then
            CommonDialog1.ShowSave
            If CommonDialog1.Filename = "" Then Exit Sub
            fnum = FreeFile
            Open CommonDialog1.Filename For Output As fnum
            Print #fnum, clist
            Close #fnum
            MsgBox "Error log saved succesfully to: " + CommonDialog1.Filename
            
    End If
    pr1.Value = 0
Exit Sub
errtrap:
MsgBox "An unexpected error has accured make sure the disk is not full or write protected."
Exit Sub


End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuexpdf_Click()
On Error GoTo errtrap
' type of copy = 1 = clipboard
' type of copy = 2 = file
Dim clist As String
Dim ff As String
Dim oldg As Long
Form5.Show 1
If BPressed = 2 Then Exit Sub
Dim i As Long
Dim fnum As Long
oldg = -1
Label1.Caption = "Building list"
pr1.Max = 100
For i = 0 To mDupCount

pr1.Value = (100 / (mDupCount)) * i
If CP = True Then DoEvents
    If FNOnly = True Then
        ff = GetFileName(mDuplicate(i).Filename)
    Else
        ff = mDuplicate(i).Filename
    End If
    
If mDuplicate(i).DuplicateNum <> oldg Then
    oldg = mDuplicate(i).DuplicateNum
    ff = "Duplicate Group: " + Trim(Str(mDuplicate(i).DuplicateNum)) + " File size: " + Trim(Str(mDuplicate(i).FileSize)) + " Bytes" + vbCrLf + "    " + ff
clist = clist + vbCrLf + ff
Else

  clist = clist + vbCrLf + "    " + ff
  End If
  If CP = True Then DoEvents
Next
    ' if it is to export to clipboard
    If TypeOfCopy = 1 Then Clipboard.Clear: Clipboard.SetText clist: MsgBox "Duplicate files found list copied to clipboard succesfully"
    ' if it is to export to file
    If TypeOfCopy = 2 Then
            CommonDialog1.ShowSave
            If CommonDialog1.Filename = "" Then Exit Sub
            fnum = FreeFile
            Open CommonDialog1.Filename For Output As fnum
            Print #fnum, clist
            Close #fnum
            MsgBox "Duplicate files list saved succesfully to: " + CommonDialog1.Filename
            
    End If
    pr1.Value = 0
Exit Sub
errtrap:
MsgBox "An unexpected error has accured make sure the disk is not full or write protected."
Exit Sub

End Sub

Private Sub mnuexpff_Click()
On Error GoTo errtrap
' type of copy = 1 = clipboard
' type of copy = 2 = file
Dim clist As String
Dim ff As String
Dim i As Long
Dim fnum As Long
Form5.Show 1
If BPressed = 2 Then Exit Sub

    pr1.Max = 100
For i = 0 To mSearchCount
    If FNOnly = True Then
        ff = GetFileName(mSearchRs(i).Filename)
    Else
        ff = mSearchRs(i).Filename
    End If
pr1.Value = (100 / mSearchCount) * i
        clist = clist + vbCrLf + ff
        If CP = True Then DoEvents
Next
    ' if it is to export to clipboard
    If TypeOfCopy = 1 Then Clipboard.Clear: Clipboard.SetText clist: MsgBox "Files found list copied to clipboard succesfully"
    ' if it is to export to file
    If TypeOfCopy = 2 Then
            CommonDialog1.ShowSave
            If CommonDialog1.Filename = "" Then Exit Sub
            fnum = FreeFile
            Open CommonDialog1.Filename For Output As fnum
            Print #fnum, clist
            Close #fnum
            MsgBox "Files found list saved succesfully to: " + CommonDialog1.Filename
            
    End If
    pr1.Value = 0
Exit Sub
errtrap:
MsgBox "An unexpected error has accured make sure the disk is not full or write protected."
Exit Sub

End Sub

Private Sub mnumovnon_Click()
NonDupAction 2
End Sub

Private Sub mnunond_Click()
On Error GoTo errtrap
' type of copy = 1 = clipboard
' type of copy = 2 = file
Dim clist As String
Dim ff As String
Dim i As Long
Dim fnum As Long
Form5.Show 1
pr1.Max = 100
If BPressed = 2 Then Exit Sub

    
For i = 0 To mSearchCount
If mSearchRs(i).IsDuplicate = False Then
    If FNOnly = True Then
        ff = GetFileName(mSearchRs(i).Filename)
    Else
        ff = mSearchRs(i).Filename
    End If
    
    clist = clist + vbCrLf + ff
Else
ff = ""
End If
pr1.Value = (100 / mSearchCount) * i
If CP = True Then DoEvents
Next
    ' if it is to export to clipboard
    If TypeOfCopy = 1 Then Clipboard.Clear: Clipboard.SetText clist: MsgBox "Non-Duplicate files list copied to clipboard succesfully"
    ' if it is to export to file
    If TypeOfCopy = 2 Then
            CommonDialog1.ShowSave
            If CommonDialog1.Filename = "" Then Exit Sub
            fnum = FreeFile
            Open CommonDialog1.Filename For Output As fnum
            Print #fnum, clist
            Close #fnum
            MsgBox "Non-Duplicate file list saved succesfully to: " + CommonDialog1.Filename
            
    End If
    pr1.Value = 0
Exit Sub
errtrap:
MsgBox "An unexpected error has accured make sure the disk is not full or write protected."
Exit Sub

End Sub

Private Sub mnuOOE_Click()
Dim CopyCount As Long
Dim i As Long
Dim lastg As Long
Dim errstring As String
CopyCount = -1
lastg = -1

For i = 0 To mSearchCount
If mSearchRs(i).IsDuplicate <> True Then
CopyCount = CopyCount + 1
mCopyFile(CopyCount).mFileName = mSearchRs(i).Filename

End If
If CP = True Then DoEvents
Next
For i = 0 To mDupCount
If CP = True Then DoEvents
If mDuplicate(i).DuplicateNum <> lastg Then
lastg = mDuplicate(i).DuplicateNum
CopyCount = CopyCount + 1
mCopyFile(CopyCount).mFileName = mDuplicate(i).Filename
End If
Next
On Error GoTo errtrap
Dim PathToCopy As String
Dim ind, CopyErrorCount As Long
Dim CopyFile As String
Dim CopyWholeFileName As String
Dim CopyFileEXT As String
Dim CopyFileName As String
Dim EndFileName As String
Dim strIND
PathToCopy = Browse(Me.hwnd, "Select destination folder")
If PathToCopy = "" Then Exit Sub

PathToCopy = AddSlash(PathToCopy)
pr1.Max = 100

For i = 0 To CopyCount
pr1.Value = (100 / CopyCount) * i

CopyFile = mCopyFile(i).mFileName
    CopyWholeFileName = GetFileName(CopyFile)
         GetFileEXT CopyWholeFileName, CopyFileName, CopyFileEXT

If CopyFile = "" Then GoTo mNext
Retry:
If ind <> 0 Then strIND = "(" + "(" + Str(ind) + ")" + ")" Else strIND = ""
EndFileName = PathToCopy + " " + CopyFileName + strIND + CopyFileEXT
If Dir(EndFileName) Then ind = ind + 1: GoTo Retry Else ind = 0
FileCopy CopyFile, EndFileName


mNext:
'temp = GetTickCount()
DoEvents

'aft = GetTickCount()
'd = d + (aft - temp)


Next
Dim errmes As String
errstring = "Errors: %"
MsgBox Replace(errstring, "%", Str(CopyErrorCount)) + errmes
pr1.Value = 0
Exit Sub
errtrap:

CopyErrorCount = CopyErrorCount + 1
errmes = errmes + vbCrLf + Err.Description
Err.Clear
Resume Next
End Sub

Private Sub mnuopt_Click()
Load Form4
Form4.Text1 = Trim(Str(Maxr))
Form4.Text2 = File1.Pattern
Form4.Text3 = Trim(Str(NOB))
If CP = True Then
Form4.Check1.Value = vbChecked
Else
Form4.Check1.Value = vbUnchecked
End If
If ZBFAD = True Then Form4.Check2.Value = vbChecked Else Form4.Check2.Value = vbUnchecked
If AskPerm = True Then Form4.Check3.Value = vbChecked Else Form4.Check3.Value = vbUnchecked
If File1.Hidden = True Then Form4.Check4.Value = vbChecked Else Form4.Check4.Value = vbUnchecked
If File1.ReadOnly = True Then Form4.Check5.Value = vbChecked Else Form4.Check5.Value = vbUnchecked
If File1.System = True Then Form4.Check6.Value = vbChecked Else Form4.Check6.Value = vbUnchecked
Form4.Show 1
End Sub

Private Sub mnupause_Click()
PauseSearch = True
End Sub

Private Sub mnusss_Click()
MsgBox "Search Complete " + vbCrLf + "Summary: " + vbCrLf + _
 Str(mSearchCount + 1) + " Matches " + vbCrLf + Str(mDupCount + 1) _
 + " Duplicates Found In " + GetTimeFromSeconds(ttc) + vbCrLf + Str(mDupGroupCount + 1) + " Duplicate Groups" + vbCrLf + Str(ErrorCount + 1) + " Errors "
End Sub
Public Sub Sort(ByVal first As Long, ByVal last As Long, d As Long)

Dim r As Long
Dim i As Long
Dim f As Long
    Dim pivot As Double
  
    
    
    If first < last Then
        pivot = Partition(first, last, d)
        Sort first, pivot - 1, d
        Sort pivot + 1, last, d
    End If
        f = GetTickCount()
            r = (f - d) / 1000
            Label10.Caption = GetTimeFromSeconds(r)
End Sub

Public Function Partition(first As Long, ByVal last As Long, d As Long) As Double
Dim f As Long
        Dim r As Long
    Dim up As Long
    Dim down As Long
    Dim pivot As Single
        f = GetTickCount()
            r = (f - d) / 1000
            Label10.Caption = GetTimeFromSeconds(r)
    pivot = mSearchRs(first).FileSize
    up = first
    down = last
    
  
    Do While (up < down)
If CP = True Then DoEvents
       
        Do While (mSearchRs(up).FileSize <= pivot) And (up < last)
            up = up + 1
        Loop
        
        Do While (mSearchRs(down).FileSize > pivot)
            down = down - 1
        Loop
        If up < down Then Swap up, down
    Loop
    
    Swap first, down
    Partition = down
End Function

Private Sub mnustop_Click()
StopSearch = True
End Sub

Private Sub mnustrt_Click()
Command1_Click

End Sub

Private Sub mnuviewdup_Click()
DuplicateAction 3
End Sub

Private Sub mnuviewnon_Click()
NonDupAction 3
End Sub

Private Sub movdup_Click()
DuplicateAction 2
End Sub

Private Sub much_Click()
Load Form6
Dim i As Integer
For i = 0 To SFC
Form6.List1.AddItem SearchFolder(i).FolderPath
Form6.List1.Selected(i) = SearchFolder(i).Used
Next
For i = 0 To BFC
Form6.List2.AddItem BannedFolder(i).FolderPath
Form6.List2.Selected(i) = BannedFolder(i).Used
Next
If SSD Then Form6.Check1.Value = vbChecked Else Form6.Check1.Value = vbUnchecked
Form6.Show 1

End Sub

