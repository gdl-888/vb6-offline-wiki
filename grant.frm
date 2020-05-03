VERSION 5.00
Begin VB.Form Dialog1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "부여"
   ClientHeight    =   1500
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "grant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox loginhis 
      Caption         =   "loginhis"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox grant 
      Caption         =   "grant"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox dev 
      Caption         =   "dev"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox ban 
      Caption         =   "ban"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox admin 
      Caption         =   "admin"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   3135
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
   Begin VB.Label Label1 
      Caption         =   "이름:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Dialog1"
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

Private Sub dev_Click()
    If Form1.username.Caption <> "개발자" And dev.Value = 1 Then
        dev.Value = 0
        MsgBox "개발자만 가능합니다.", 16, "오류"
    End If
End Sub

Private Sub grant_Click()
    If Text1.Text = "개발자" And grant.Value = 0 Then
    grant.Value = 1
    MsgBox "그렇게 부여하면 안됩니다.", 16, "오류"
    End If
End Sub

Private Sub OKButton_Click()
    On Error Resume Next
    
    If FileExists("C:\wiki\users\" & Text1.Text) = False Then
        MsgBox "없는 사용자.", 16, "부여"
        Unload Me
    End If
    
    If Text1.Text = "개발자" Then
        If grant.Value = 0 Or dev.Value = 0 Then
            MsgBox "그렇게 부여하면 안됩니다.", 16, "오류"
            Unload Me
        End If
    End If
    
    Dim iFileNo As Integer
    
    If admin.Value = Checked Then
        iFileNo = FreeFile
        'open the file for writing
        Open "C:\WIKI\GRANT\admin\" & Text1.Text For Output As #iFileNo
           'please note, if this file already exists it will be overwritten!
           'write some example text to the file
           
           'close the file (if you dont do this, you wont be able to open it again!)
        Close #iFileNo
    Else
        Kill ("C:\WIKI\GRANT\admin\" & Text1.Text)
    End If
    
    
    
    If ban.Value = Checked Then
        iFileNo = FreeFile
        'open the file for writing
        Open "C:\WIKI\GRANT\ban\" & Text1.Text For Output As #iFileNo
           'please note, if this file already exists it will be overwritten!
           'write some example text to the file
           
           'close the file (if you dont do this, you wont be able to open it again!)
        Close #iFileNo
    Else
        Kill ("C:\WIKI\GRANT\ban\" & Text1.Text)
    End If
    
    If dev.Value = Checked Then
        iFileNo = FreeFile
        'open the file for writing
        Open "C:\WIKI\GRANT\dev\" & Text1.Text For Output As #iFileNo
           'please note, if this file already exists it will be overwritten!
           'write some example text to the file
           
           'close the file (if you dont do this, you wont be able to open it again!)
        Close #iFileNo
    Else
        Kill ("C:\WIKI\GRANT\dev\" & Text1.Text)
    End If
    
    If grant.Value = Checked Then
        iFileNo = FreeFile
        'open the file for writing
        Open "C:\WIKI\GRANT\grant\" & Text1.Text For Output As #iFileNo
           'please note, if this file already exists it will be overwritten!
           'write some example text to the file
           
           'close the file (if you dont do this, you wont be able to open it again!)
        Close #iFileNo
    Else
        Kill ("C:\WIKI\GRANT\grant\" & Text1.Text)
    End If
    
    If loginhis.Value = Checked Then
        iFileNo = FreeFile
        'open the file for writing
        Open "C:\WIKI\GRANT\loginhis\" & Text1.Text For Output As #iFileNo
           'please note, if this file already exists it will be overwritten!
           'write some example text to the file
           
           'close the file (if you dont do this, you wont be able to open it again!)
        Close #iFileNo
    Else
        Kill ("C:\WIKI\GRANT\loginhis\" & Text1.Text)
    End If
    
    
    
    Unload Me
End Sub
