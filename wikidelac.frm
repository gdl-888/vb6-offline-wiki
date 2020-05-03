VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  '단일 고정
   Caption         =   "계정 강제 삭제"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "유저 이름:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valid As Boolean

Private Sub Command1_Click()
    On Error Resume Next
    If MsgBox("정말로?", vbYesNo + vbQuestion, "계정 강제 삭제") = vbYes Then
        On Error Resume Next
        If Text1.Text = "개발자" Then
            MsgBox "그렇게 하면 안됩니다.", 16, "삭제"
        Else
            Kill ("C:\WIKI\USERS\" & Text1.Text)
            Unload Me
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
    Form1.Show
End Sub
