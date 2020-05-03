VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "사용자 목록"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2535
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "닫기"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   2250
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    File1.Path = "c:\wiki\users"
End Sub
