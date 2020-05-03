VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "삭제..."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "R.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Line Line3 
      X1              =   960
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   840
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFCC00&
      BackStyle       =   1  '투명하지 않음
      Height          =   975
      Left            =   0
      Shape           =   3  '원형
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "삭제: C:\WIKI"
      Height          =   240
      Left            =   375
      TabIndex        =   0
      Top             =   1035
      Width           =   3135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    WindowsMediaPlayer1.settings.setMode "loop", True
End Sub

