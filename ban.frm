VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "차단"
   ClientHeight    =   2685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "ban.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "해제"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "차단"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   480
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
   Begin VB.Label Label2 
      Caption         =   "메모:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "유저 이름:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Dialog2"
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

Private Sub Form_Load()
    Option1.Value = True
End Sub

Private Sub OKButton_Click()
On Error Resume Next

    If FileExists("C:\wiki\users\" & Text1.Text) = False Then
        MsgBox "없는 사용자.", 16, "차단"
        Unload Me
    End If
    
    If Option1.Value = True Then
    If Text1.Text = "개발자" Then
        MsgBox "그렇게 하면 안됩니다.", 16, "오류"
    Else
    Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\BLOCK\" & Text1.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       Print #iFileNo, Text2.Text
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      End If
    Else
        Kill ("C:\WIKI\BLOCK\" & Text1.Text)
    End If
    Unload Me
End Sub
