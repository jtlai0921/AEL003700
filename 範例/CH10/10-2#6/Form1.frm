VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3165
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   372
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
   Begin VB.ListBox List1 
      Height          =   1140
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0019
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
      Height          =   252
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "��ܿ��������:"
      Height          =   252
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   For i = 0 To List1.ListCount - 1
      If List1.Selected(i) Then   '�p�G��i�����Q���
         Label2 = List1.List(i)   '�N�N�䤺�e��ܦbLabel2
         Exit For                 '�ø����j��
      End If
   Next i
End Sub


