VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton command2 
      Caption         =   "����"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton command1 
      Caption         =   "��X"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "��X���G:"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�m�W:"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   Label3.Caption = "�w��" + Text1.Text + "�Ӿ�Viaual BASIC"
End Sub

Private Sub command2_Click()
   End
End Sub


