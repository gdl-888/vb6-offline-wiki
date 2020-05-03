VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "구성"
   ClientHeight    =   2400
   ClientLeft      =   2640
   ClientTop       =   2430
   ClientWidth     =   6645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "취소"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(Θ)"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   8
      Text            =   "1234"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Text            =   "C:\WIKI"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Text            =   "개발자"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "소유자       비밀번호:"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "데이터   디렉토리:"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "소유자계정명:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "위키 이름:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    If Text2.Text <> "개발자" Then
        MsgBox "소유자 계정명이 '개발자'가 아닐 경우 악성 사용자에게 불이익을 당할 수 있습니다." & vbNewLine & " - 개발자 권한을 부여 또는 회수할 수 없습니다." & vbNewLine & " - 악성 관리자가 당신의 계정을 차단할 수 있습니다." & vbNewLine & " - 다른 사무관이 당신의 사무관 권한을 훔쳐갈 수 있습니다." & vbNewLine & " - 개발자 권한을 가진 다른 사용자가 당신의 계정을 파괴할 수 있습니다.", vbOKCancel + vbExclamation, "경고"
    End If
    Form9.Show
    MkDir "C:\WIKI\"
    MkDir "C:\wiki\users\"
    MkDir "C:\wiki\block\"
    MkDir "C:\wiki\changes\"
    MkDir "C:\wiki\discuss\"
    MkDir "C:\wiki\history\"
    MkDir "C:\wiki\doc\"
    MkDir "C:\wiki\acl\"
    MkDir "C:\wiki\acl\read\"
    MkDir "C:\wiki\acl\edit\"
    MkDir "C:\wiki\acl\discuss\"
    MkDir "C:\wiki\acl\acl\"
    MkDir "C:\wiki\acl\request\"
    MkDir "C:\wiki\grant\"
    MkDir "C:\wiki\grant\loginhis\"
    MkDir "C:\wiki\grant\grant\"
    MkDir "C:\wiki\grant\ban\"
    MkDir "C:\wiki\grant\admin\"
    MkDir "C:\wiki\grant\dev\"
    MkDir "C:\wiki\loginhis\"
    MkDir "C:\wiki\discuss\close\"
    MkDir "C:\wiki\block\date\"
    MkDir "C:\WIKI\rev\"
    MkDir "C:\WIKI\request\"
    Dim cnt As Integer
    Unload Form9
    Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\GRANT\loginhis\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\GRANT\grant\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\GRANT\ban\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\GRANT\admin\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\GRANT\dev\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      
      
      
      
    If Text1.Text <> "" And Text2.Text <> "" Then
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\WIKINAME.TXT" For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text1.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      
      iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\USERS\" & Text2.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text4.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      
      
    Dim fh As Integer ' file handle
Dim MyLine As String 'a single line from the file
fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\WIKINAME.txt" For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
Form1.Caption = MyLine & "@대문" & " - " & MyLine
wikiname = MyLine
Wend
Close #fh 'close the file so someone else can read it
doctitle = wikiname & "@대문"
    Form1.wiki1.Caption = wikiname
    
    Form1.Show
    Unload Me
    Else
    MsgBox "위키 이름의 값은 필수입니다.", 16, "구성"
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Caption = "(Θ)" Then
        Command2.Caption = "(Ο)"
        Text4.PasswordChar = ""
    Else
        Command2.Caption = "(Θ)"
        Text4.PasswordChar = "*"
    End If
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
    Form9.Show
End Sub
