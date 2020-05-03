VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form10 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "삭제..."
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   570
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3705
      URL             =   "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Videos\FILEDELR.AVI"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6535
      _cy             =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "삭제: C:\WIKI"
      Height          =   240
      Left            =   375
      TabIndex        =   0
      Top             =   1035
      Width           =   3135
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WindowsMediaPlayer1.settings.setMode "loop", True
End Sub

