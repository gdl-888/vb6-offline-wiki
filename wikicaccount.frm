VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "���� �����"
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
      IMEMode         =   3  '��� ����
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  '��� ����
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
      Caption         =   "���"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   " * ������ Ż��� �Ұ����մϴ�."
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ȣ Ȯ��:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "��ȣ:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "����� ID:"
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
        If FileExists("C:\wiki\users\" & Text1.Text) = False Or Text1.Text <> "�͸�" Then
            Dim iFileNo As Integer
        iFileNo = FreeFile
       'open the file for writing
      Open "C:\WIKI\USERS\" & Text1.Text For Output As #iFileNo
       'please note, if this file already exists it will be overwritten!
       
       'write some example text to the file
      Print #iFileNo, Text2.Text
       
       'close the file (if you dont do this, you wont be able to open it again!)
      Close #iFileNo
      MsgBox "ȯ���մϴ�! " & Text1.Text & "���� ���� ������ �Ϸ�Ǿ����ϴ�.", vbInformation, "���� �����"
      Unload Me
      Else
      MsgBox "�̹� �ִ� ������Դϴ�.", 16, "���� �����"
      End If
      Else
      MsgBox "��ȣ�� �ٽ� Ȯ���ϼ���.", 16, "���� �����"
      End If
      Else
      MsgBox "����� �̸� �Ǵ� ��ȣ�� ������ ���� �°� �Է����ּ���.", 16, "���� �����"
    End If
End Sub
