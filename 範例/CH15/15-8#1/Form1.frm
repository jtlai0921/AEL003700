VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4155
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton cmdQuery 
      Caption         =   "�d��"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "�s�W"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
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
      Left            =   3000
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "�}��"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtEng 
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
      Left            =   1440
      TabIndex        =   6
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtChin 
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
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtName 
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
      Left            =   1440
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "�^�妨�Z:"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "��妨�Z:"
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "�m        �W:"
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
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '��u�T�w
      Caption         =   " �� �� �� �Z �n �� �@ �~"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
   Open "a:\score.dat" For Output As #1
   Close #1
End Sub

Private Sub cmdAdd_Click()
   Open "a:\score.dat" For Append As #1
   Write #1, txtName, txtChin, txtEng
   Close #1
   txtName = "": txtChin = "": txtEng = ""
End Sub


Private Sub cmdQuery_Click()
   Open "a:\score.dat" For Input As #1
   Dim find As Boolean
   find = False
   Do While Not EOF(1)
      Input #1, nam, chin, eng
      If nam = txtName Then
         find = True
         txtChin = chin
         txtEng = eng
         Exit Do
     End If
   Loop
   If find Then
      MsgBox "�����,�äw���!"
   Else
      MsgBox "�ɮפ��S���n�䪺�m�W:" + txtName
   End If
   Close #1
   txtName = "": txtChin = "": txtEng = ""
End Sub

Private Sub cmdEnd_Click()
   End
End Sub


