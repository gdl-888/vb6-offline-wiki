VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FF00FF&
   Caption         =   "����"
   ClientHeight    =   1515
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   1515
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Line Line3 
      X1              =   960
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   840
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      X1              =   480
      X2              =   480
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFCC00&
      BackStyle       =   1  '�������� ����
      Height          =   975
      Left            =   0
      Shape           =   3  '����
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "����: C:\WIKI"
      Height          =   240
      Left            =   630
      TabIndex        =   0
      Top             =   1110
      Width           =   3135
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
