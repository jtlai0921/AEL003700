VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�D�̤j���]�ƻP�̤p������"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4020
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "�p��"
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
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtno2 
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
      Left            =   2640
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtno1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Text            =   " "
      Top             =   225
      Width           =   615
   End
   Begin VB.Label lblLCM 
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
      DataField       =   " "
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblGCD 
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
      DataField       =   " "
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "�̤p�����Ƭ�:"
      DataField       =   " "
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
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "�̤j���]�Ƭ�:"
      DataField       =   " "
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
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "�п�J�ĤG�Ӿ��:"
      DataField       =   " "
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�Ĥ@�Ӿ��:"
      DataField       =   " "
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
  Do   '�ˬd�Ĥ@�ӼƬO�_�������,���O�N�n�D���s��J
    m = Val(txtno1)
    If m > 0 And m = Int(m) Then Exit Do
    MsgBox "�Ĥ@�ӼƫD�����,�Э��s��J��A���p��s"
    txtno1 = ""           '�N��J���Ĥ@�ӼƲM���ť�
    txtno1.SetFocus  '�]�w�n�I����,�Y���J�I���b����
    Exit Sub         '�������{��
  Loop
  Do   '�ˬd�ĤG�ӼƬO�_�������,���O�N�n�D���s��J
    n = Val(txtno2)
    If n > 0 And n = Int(n) Then Exit Do
    MsgBox "�ĤG�ӼƫD�����,�Э��s��J��A���p��s"
    txtno2 = ""
    txtno2.SetFocus
    Exit Sub
  Loop
  If m > n Then k = n Else k = m   '�]�wm�Pn�����p��
  Rem ��X��ƪ��̤j���]��
  For i = k To 1 Step -1
    If m / i = Int(m / i) And n / i = Int(n / i) Then Exit For  '���F
  Next i
  lblGCD = i                '��̤ܳj���]��
  lblLCM = m * n / i        '��̤ܳp������
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

