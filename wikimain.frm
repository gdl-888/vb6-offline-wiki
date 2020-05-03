VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   ClientHeight    =   5310
   ClientLeft      =   930
   ClientTop       =   1245
   ClientWidth     =   7125
   Icon            =   "wikimain.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  '사용자 정의
   ScaleHeight     =   5310
   ScaleWidth      =   7125
   Begin VB.Timer chack 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6855
      Begin VB.CommandButton Command5 
         Caption         =   "새 문서 만들기(&O)"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "해당 문서를 찾을 수 없습니다."
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin VB.CommandButton search 
      Caption         =   "→"
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox stext 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin MSForms.MultiPage MultiPage1 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "wikimain.frx":030A
      TabIndex        =   6
      Top             =   480
      Width           =   6855
   End
   Begin VB.Menu wiki1 
      Caption         =   "테스트"
      Begin VB.Menu home 
         Caption         =   "대문(&H)"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu vb6explnk 
         Caption         =   "VB6 실험실"
         Visible         =   0   'False
      End
      Begin VB.Menu gdfryhhy 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu recent 
      Caption         =   "최근 변경"
      Begin VB.Menu rchanges 
         Caption         =   "최근 변경(&C)"
      End
      Begin VB.Menu rdiscuss 
         Caption         =   "최근 토론(&D)"
      End
   End
   Begin VB.Menu special 
      Caption         =   "특수 기능"
      Begin VB.Menu checkbl 
         Caption         =   "차단 여부 확인(&K)"
      End
      Begin VB.Menu userlsty 
         Caption         =   "사용자 목록(&U)"
      End
      Begin VB.Menu grantusr 
         Caption         =   "권한 사용자 목록(&G)"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu saccount 
         Caption         =   "사용자 차단(&U)"
         Visible         =   0   'False
      End
      Begin VB.Menu granta 
         Caption         =   "권한 부여(&G)"
         Visible         =   0   'False
      End
      Begin VB.Menu loghis 
         Caption         =   "로그인 내역(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu cwname 
         Caption         =   "위키 이름 변경(&K)"
         Visible         =   0   'False
      End
      Begin VB.Menu raccnt 
         Caption         =   "계정 강제 삭제(&Q)"
         Visible         =   0   'False
      End
      Begin VB.Menu reset 
         Caption         =   "위키 초기화(&Y)"
         Visible         =   0   'False
      End
      Begin VB.Menu cpass 
         Caption         =   "사용자 암호 보기(&P)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu username 
      Caption         =   "익명"
      NegotiatePosition=   3  '오른쪽
      Begin VB.Menu countri 
         Caption         =   "기여(&O)"
      End
      Begin VB.Menu login 
         Caption         =   "로그인(&L)"
      End
      Begin VB.Menu logou 
         Caption         =   "로그아웃(&L)"
         Visible         =   0   'False
      End
      Begin VB.Menu chpasf 
         Caption         =   "암호 변경(&S)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In a standard Module: Module1.bas

    Dim doctitle As String
    Dim wikiname As String
    Dim docmode As Integer

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

Function TitleSet()
    Me.Caption = doctitle & " - " & wikiname
End Function


Private Sub chack_Timer()
    If FileExists("C:\WIKI\block\" & username.Caption) = True Then
        Dim fh As Integer
        Dim MyLine As String
        fh = FreeFile 'get a free file handle from the OS
        Open "C:\WIKI\block\" & username.Caption For Input As #fh 'Open the file for reading
        While Not EOF(fh) 'are we at the End Of the File
        Line Input #fh, MyLine 'actually read a line from the file
        MsgBox "현재 접속된 계정은 차단된 것으로 확인됩니다(" & MyLine & ")." & vbCrLf & vbCrLf & "지금 나갑니다.", 16, "로그인"
        Wend
        Close #fh 'close the file so someone else can read it
        logou_Click
    End If
End Sub

Private Sub checkbl_Click()
    Dim uname As String
    uname = InputBox("유저 이름:", "차단 여부 확인")
    If FileExists("C:\wiki\users\" & uname) = True Then
    If FileExists("c:\wiki\block\" & uname) = False Then
        MsgBox "차단되지 않았습니다.", vbInformation, "차단 여부"
    Else
        MsgBox "차단되어 있습니다.", vbInformation, "차단 여부"
    End If
    Else
    MsgBox "없는 사용자입니다.", vbCritical, "알림"
    End If
End Sub

Private Sub chpasf_Click()
    If username.Caption = "사용자" Or username.Caption = "익명" Then
        MsgBox "특수 계정의 암호를 변경할 수 없습니다.", 16, "오류"
    Else
        Dim userng As String
        userng = username.Caption
      Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\users\" & userng For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       Print #iFileNo, InputBox("새 암호: ", "암호 변경")
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      MsgBox "완료.", vbInformation, "알림"
    End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
      Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\DOC\" & doctitle For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
    Form3.Caption = doctitle & " (편집) - " & wikiname
    Form3.Show
    Me.Hide
End Sub

Private Sub cpass_Click()
    Dim fh As Integer ' file handle
Dim MyLine As String 'a single line from the file
    Dim cpuname As String
    cpuname = InputBox("유저 이름: ", "암호 확인")
    If FileExists("c:\wiki\users\" & cpuname) = True Then
        fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\users\" & cpuname For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
MsgBox cpuname & "의 암호: " & MyLine, vbInformation, "암호"
wikiname = MyLine
Wend
Close #fh 'close the file so someone else can read it
    Else
        MsgBox "없는 사용자입니다.", 16, "오류"
    End If
End Sub

Private Sub cwname_Click()
    Me.Hide
    Form5.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If FileExists("C:\WIKI\WIKINAME.TXT") = False Then
        Form2.Show
        Me.Hide
    Else
    Dim fh As Integer ' file handle
Dim MyLine As String 'a single line from the file
fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\WIKINAME.txt" For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
Me.Caption = MyLine & "@대문" & " - " & MyLine
wikiname = MyLine
Wend
Close #fh 'close the file so someone else can read it
doctitle = wikiname & "@대문"
    wiki1.Caption = wikiname
    End If
    
    
    
    If FileExists("C:\WIKI\DOC\" & doctitle) = False Then
        Label1.Visible = True
        Command5.Visible = True
        Label2.Visible = False
    Else
        Label1.Visible = False
        Command5.Visible = False
        Label2.Visible = True
        fh = FreeFile 'get a free file handle from the OS
Open "C:\WIKI\DOC\" & doctitle For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
Label2.Caption = MyLine
Wend
Close #fh 'close the file so someone else can read it
    End If
End Sub

Private Sub Form_Resize()
    search.Left = Me.Width - 720
    stext.Width = Me.Width - 945
    Frame1.Height = Me.Height - 1830
    Frame1.Width = Me.Width - 390
    MultiPage1.Width = Me.Width - 390
End Sub

Private Sub granta_Click()
    Dialog1.Show
End Sub

Private Sub grantusr_Click()
    Form8.Show
End Sub

Private Sub home_Click()
    doctitle = wikiname & ":대문"
End Sub

Private Sub loghis_Click()
    Dim usern As String
    usern = InputBox("유저 이름: ", "로그인 내역")
    If FileExists("c:\wiki\loginhis\" & usern) = False Then
        MsgBox "없거나 한 번도 로그인하지 않은 사람입니다.", 16, "로그인 내역"
    Else
        Dim fh As Integer ' file handle
        Dim MyLine As String 'a single line from the file
        fh = FreeFile 'get a free file handle from the OS
        Open "C:\WIKI\loginhis\" & usern For Input As #fh 'Open the file for reading
        While Not EOF(fh) 'are we at the End Of the File
        Line Input #fh, MyLine 'actually read a line from the file
        MsgBox "로그인 횟수: " & MyLine, vbInformation, "로그인 내역"
        Wend
        Close #fh 'close the file so someone else can read it
    End If
End Sub

Private Sub login_Click()
    frmLogin.Show
End Sub

Private Sub logou_Click()
On Error Resume Next
    username.Caption = "익명"
    logou.Visible = False
    login.Visible = True
    dash2.Visible = False
    saccount.Visible = False
    granta.Visible = False
    loghis.Visible = False
    dash3.Visible = False
    raccnt.Visible = False
    cwname.Visible = False
    reset.Visible = False
    cpass.Visible = False
End Sub

Private Sub MultiPage1_Change()
    docmode = MultiPage1.Value + 1
    On Error Resume Next
    If docmode = 2 Then
    If FileExists("c:\wiki\doc\" & doctitle) = False Then
        Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\DOC\" & doctitle For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      
    End If
    
    Dim fh As Integer ' file handle
Dim MyLine As String 'a single line from the file
fh = FreeFile 'get a free file handle from the OS
Open "c:\wiki\doc\" & doctitle For Input As #fh 'Open the file for reading
While Not EOF(fh) 'are we at the End Of the File
Line Input #fh, MyLine 'actually read a line from the file
Form3.Text1.Text = MyLine
Wend
Close #fh 'close the file so someone else can read it
    
    Form3.Caption = doctitle & " (편집) - " & wikiname
    Form3.Show
    Form3.Label2.Caption = doctitle
        Me.Hide
        
    End If
    'MultiPage1.Value = 1
    
    If docmode = 4 Then
        If MsgBox("정말로", vbQuestion + vbYesNo, "삭제") = vbOK Then
            Kill ("C:\WIKI\DOC\" & doctitle)
        End If
    End If
End Sub

Private Sub raccnt_Click()
    Form4.Show
    Me.Hide
End Sub

Private Sub reset_Click()
    If MsgBox("이 작업은 영구적이며 되돌릴 수 없습니다. 계속하시겠습니까?", vbYesNo + vbQuestion, "위키 초기화") = vbYes Then
        If MsgBox("위키의 모든 문서, 토론, 사용자를 영구적으로 삭제합니다. 정말 삭제하시겠습니까?", vbYesNo + vbExclamation, "위키 초기화") = vbYes Then
            Form6.Show
            Me.Hide
        End If
    End If
End Sub

Private Sub saccount_Click()
    Dialog2.Show
End Sub

Private Sub userlsty_Click()
    Form7.Show
End Sub

