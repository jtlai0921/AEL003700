VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4230
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   1
      Left            =   2400
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Picture1(0).Visible = False Then    '���p����Ϭݤ����N�K
     Picture1(0).Picture = Picture1(1).Picture  '�N�k��ϳ]�w��������
     Picture1(0).Visible = True           '�]�w����Ϭݱo��
     Picture1(1).Visible = False          '�]�w�k��Ϭݤ���
   Else                            '����Ϭݱo�������p
     Picture1(1).Picture = Picture1(0).Picture  '�N����ϳ]�w���k����
     Picture1(1).Visible = True           '�]�w�k��Ϭݱo��
     Picture1(0).Visible = False          '�]�w����Ϭݤ���
  End If
End Sub
