VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�M�������򥻾ާ@"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4560
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton cmdEnd 
      Caption         =   "����"
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
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���M"
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
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�R��"
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
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.ListBox lstData 
      Height          =   1320
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdInput 
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
   Begin VB.TextBox txtInput 
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
      Caption         =   "�w��J����ƶ�:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�n�s�W�����:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInput_Click()
   newdata = txtInput.Text   '�N��r��������e�s�i�ܼ�newdata��
   For i = 0 To lstData.ListCount - 1
      If lstData.Selected(i) Then   '���p��i���Q����N
         lstData.AddItem newdata, i '�N�ܼ�newdata�����e�[�J�M�椤
         flag = 1    '�]�w�������ƶ������p
         Exit For    '�����j��
      End If
   Next i
   If flag = 0 Then lstData.AddItem newdata '�S�������ƶ��N�[�b����
   txtInput.Text = ""      '�N��r����M���Ŧr��
   txtInput.SetFocus       '�N�n�I�]�w�b��J��ƪ���r����W
End Sub

Private Sub cmdDelete_Click()
   For i = 0 To lstData.ListCount - 1
      If lstData.Selected(i) Then   '���p��i���Q����N
         lstData.RemoveItem i       '�R����i��
         Exit For                    '�����j��
      End If
   Next i
End Sub

Private Sub cmdClear_Click()
   lstData.Clear        '�N�M���������e�����M����
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

