VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "��ܵ��G"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   2445
   ScaleWidth      =   3735
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "�^��J�e��"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '��u�T�w
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form2.Hide        '���ê��G
   Form1.Show        '��ܪ��@
   Form1.Text1 = ""  '�M�����@��r��������e
   Form1.Text1.SetFocus  '�N��J�I�]�w�b���@����r���
End Sub

