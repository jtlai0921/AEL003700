VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�Ȥ�
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�K�X�G"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      Print "�A��J���K�X�O�G"; Text1
End Sub

Private Sub Command2_Click()
   End
End Sub
