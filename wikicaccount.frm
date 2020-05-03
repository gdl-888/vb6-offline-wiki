VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "계정 만들기"
   ClientHeight    =   3435
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "wikicaccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text3 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  '사용 못함
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "취소"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "확인"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   " * 가입후 탈퇴는 불가능합니다."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "암호 확인:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "암호:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "사용자 ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
 
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
 
Public Function FileExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        FileExists = True
    Else
        FileExists = False
    End If
    
End Function
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
        If Text2.Text = Text3.Text Then
        If FileExists("C:\wiki\users\" & Text1.Text) = False Or Text1.Text <> "익명" Then
            Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\USERS\" & Text1.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text2.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      MsgBox "환영합니다! " & Text1.Text & "님의 계정 생성이 완료되었습니다.", vbInformation, "계정 만들기"
      Unload Me
      Else
      MsgBox "이미 있는 사용자입니다.", 16, "계정 만들기"
      End If
      Else
      MsgBox "암호를 다시 확인하세요.", 16, "계정 만들기"
      End If
      Else
      MsgBox "사용자 이름 또는 암호의 형식을 값에 맞게 입력해주세요.", 16, "계정 만들기"
    End If
End Sub
