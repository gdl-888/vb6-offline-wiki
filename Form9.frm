VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form9 
   BackColor       =   &H00FF00FF&
   Caption         =   "备己"
   ClientHeight    =   1515
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   1515
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows 扁夯蔼
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "积己: C:\WIKI"
      Height          =   240
      Left            =   630
      TabIndex        =   1
      Top             =   1110
      Width           =   3135
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   570
      Left            =   615
      TabIndex        =   0
      Top             =   315
      Width           =   3705
      URL             =   "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Videos\FILECOPY.AVI"
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
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WindowsMediaPlayer1.settings.setMode "loop", True
End Sub

