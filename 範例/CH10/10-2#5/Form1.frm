VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "�w��J���v���W�٦p�U:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�v���W��:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   ItemData = Text1.Text   '�N��r��������e�s�i�ܼ�ItemData��
   List1.AddItem ItemData  '�N�ܼ�ItemData�����e�[�J�M������
   Text1.Text = ""         '�N��r����M���Ŧr��
   Text1.SetFocus          '�N�n�I�]�w�b��J��ƪ���r����W
End Sub
