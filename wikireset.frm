VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "위키 초기화"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command22 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   4200
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "예(&Y)"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "아니요(&N)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "마지막 경고입니다. 이제 초기화하면 타임머신 따위는 없습니다. 계속하시겠습니까?"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "wikireset.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DeleteDirectory(ByVal dir_name As String)
Dim file_name As String
Dim files As Collection
Dim i As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    file_name = Dir$(dir_name & "\*.*", vbReadOnly + _
        vbHidden + vbSystem + vbDirectory)
    Do While Len(file_name) > 0
        If (file_name <> "..") And (file_name <> ".") Then
            files.Add dir_name & "\" & file_name
        End If
        file_name = Dir$()
    Loop

    ' Delete the files.
    For i = 1 To files.Count
        file_name = files(i)
        ' See if it is a directory.
        If GetAttr(file_name) And vbDirectory Then
            ' It is a directory. Delete it.
            DeleteDirectory file_name
        Else
            ' It's a file. Delete it.
          '  lblStatus.Caption = file_name
          '  lblStatus.Refresh
            SetAttr file_name, vbNormal
            Kill file_name
        End If
    Next i

    ' The directory is now empty. Delete it.
   ' lblStatus.Caption = dir_name
   ' lblStatus.Refresh

    ' Remove the read-only flag if set.
    ' (Thanks to Ralf Wolter.)
    SetAttr dir_name, vbNormal
    RmDir dir_name
End Sub

Private Sub Command1_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Command10_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command11_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command12_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command13_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command14_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command15_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command16_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command17_Click()
Form10.Show
DeleteDirectory ("C:\WIKI\")
Unload Form10
    Form2.Show
    Unload Me
End Sub

Private Sub Command18_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command19_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Command20_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command21_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command22_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command23_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command24_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command25_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command3_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Command4_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Command5_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Command6_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command7_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command8_Click()
Form1.Show
    Unload Me
End Sub

Private Sub Command9_Click()
Form1.Show
    Unload Me
End Sub
