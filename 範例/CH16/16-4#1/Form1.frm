VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��J���"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4080
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "��ܵ��G"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�m�W�G"
      BeginProperty Font 
         Name            =   "�з���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form2.Show                             '��ܪ��G
   Form2.Label1 = "HELLO!" & Form1.Text1  '�b���G��������ܵ��G
   Form1.Hide                             '���ê�棸
End Sub

