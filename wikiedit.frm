VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  '크기 고정 대화 상자
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "wikiedit.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   5640
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "저장(&S)"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5640
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label Label2 
      Height          =   135
      Left            =   600
      TabIndex        =   5
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "요약(&R):"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
    Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\DOC\" & Label2.Caption For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text1.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      
      Unload Me
    Form1.Show
End Sub

Private Sub Command2_Click()
    Unload Me
    Form1.MultiPage1.Value = 0
    Form1.Show
End Sub

