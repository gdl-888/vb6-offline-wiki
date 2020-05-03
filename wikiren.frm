VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "위키 이름 변경"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\WIKINAME.TXT" For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text1.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      MsgBox "위키 다시 시작 바람.", vbInformation, "위키 이름 변경"
      Unload Me
    Form1.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    Form1.Show
End Sub

