VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����ഫ����"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "�ഫ"
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
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtScore 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblResult 
      Alignment       =   2  '�m�����
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblGrade 
      Caption         =   "��    ��  ��"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblScore 
      Caption         =   "�п�J����:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
   Dim s As Single     '�]�wscore�������ܼ�
   s = Val(txtScore)   '�N��J��r���������ন�ƭ�
   If s >= 80 Then
      lblResult = "A"
   ElseIf s >= 70 Then
      lblResult = "B"
   ElseIf s >= 60 Then
      lblResult = "C"
   Else
      lblResult = "D"
   End If
End Sub

Private Sub cmdEnd_Click()
   End
End Sub
