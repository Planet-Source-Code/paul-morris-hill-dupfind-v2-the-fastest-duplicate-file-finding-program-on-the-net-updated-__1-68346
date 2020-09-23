VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000013&
   BorderStyle     =   0  'None
   Caption         =   "DupFind V2"
   ClientHeight    =   2670
   ClientLeft      =   4275
   ClientTop       =   3450
   ClientWidth     =   5220
   LinkTopic       =   "Form3"
   ScaleHeight     =   2670
   ScaleWidth      =   5220
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to continue"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   $"Form3.frx":0000
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "DupFind V2"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Me.Hide
End Sub

Private Sub Label1_Click()
Me.Hide

End Sub

Private Sub Label2_Click()
Me.Hide

End Sub

Private Sub Label3_Click()
Me.Hide

End Sub
