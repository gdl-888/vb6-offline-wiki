VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "권한 사용자 목록"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10920
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "닫기"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   3600
      Width           =   1575
   End
   Begin VB.FileListBox File5 
      Height          =   3150
      Left            =   8760
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.FileListBox File4 
      Height          =   3150
      Left            =   6600
      TabIndex        =   6
      Top             =   360
      Width           =   2055
   End
   Begin VB.FileListBox File3 
      Height          =   3150
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.FileListBox File2 
      Height          =   3150
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   3150
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "로그인 조회"
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "부여"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "개발자"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "차단"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "관리자"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    File1.Path = "c:\wiki\grant\admin\"
    File2.Path = "c:\wiki\grant\ban\"
    File3.Path = "c:\wiki\grant\dev\"
    File4.Path = "c:\wiki\grant\grant\"
    File5.Path = "c:\wiki\grant\loginhis\"
End Sub
