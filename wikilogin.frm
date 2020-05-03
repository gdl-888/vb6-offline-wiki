VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로그인"
   ClientHeight    =   2250
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4335
   Icon            =   "wikilogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1329.375
   ScaleMode       =   0  '사용자
   ScaleWidth      =   4070.33
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   1800
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1529
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Default         =   -1  'True
      Height          =   390
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  '사용 못함
      Left            =   1529
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "계정 만들기(&T)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "사용자 이름(&U):"
      Height          =   375
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Caption         =   "암호(&P):"
      Height          =   270
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   660
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
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

Function FileText(filename$) As String
    Dim handle As Integer
    handle = FreeFile
    Open filename$ For Input As #handle
    FileText = Input$(LOF(handle), handle)
    Close #handle
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
    Dim pass As String
    If FileExists("c:\wiki\users\" & txtUserName.Text) Then
    
        Dim fh As Integer ' file handle
Dim MyLine As String 'a single line from the file
fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\users\" & txtUserName.Text For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
If MyLine = txtPassword.Text Then
pass = MyLine
If FileExists("c:\wiki\block\" & txtUserName.Text) = True Then
fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\block\" & txtUserName For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
MsgBox "차단된 계정입니다." & vbNewLine & vbNewLine & "차단 사유: " & MyLine, 16, "로그인"
Wend
Else



Form1.username.Caption = txtUserName.Text
Form1.login.Visible = False
Form1.logou.Visible = True
If FileExists("c:\wiki\grant\ban\" & txtUserName) = True Then
    Form1.saccount.Visible = True
    Form1.dash2.Visible = True
Else
    Form1.saccount.Visible = False
End If
If FileExists("c:\wiki\grant\grant\" & txtUserName) = True Then
    Form1.granta.Visible = True
    Form1.dash2.Visible = True
Else
    Form1.granta.Visible = False
End If
If FileExists("c:\wiki\grant\loginhis\" & txtUserName) = True Then
    Form1.loghis.Visible = True
    Form1.dash2.Visible = True
Else
    Form1.loghis.Visible = False
End If
If FileExists("c:\wiki\grant\dev\" & txtUserName) = True Then
    Form1.raccnt.Visible = True
    Form1.cwname.Visible = True
    Form1.reset.Visible = True
    Form1.dash3.Visible = True
    Form1.cpass.Visible = True
Else
    Form1.dash3.Visible = False
    Form1.raccnt.Visible = False
    Form1.cwname.Visible = False
    Form1.reset.Visible = False
    Form1.cpass.Visible = False
End If
End If
Else
           If pass <> txtPassword.Text Then
                MsgBox "암호가 올바르지 않습니다.", 16, "로그인"
            End If
           Unload Me
        End If
Wend
Close #fh 'close the file so someone else can read it
            
    Else
        MsgBox "사용자 이름이 올바르지 않습니다.", 16, "로그인"
        Unload Me
    End If
    Unload Me
End Sub

Private Sub Label1_Click()
Dialog.Show
End Sub


