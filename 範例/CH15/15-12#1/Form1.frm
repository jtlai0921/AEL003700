VERSION 5.00
Begin VB.Form �H���ɪ����� 
   Caption         =   "�ǥ͸�Ƨ@�~"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   3840
   StartUpPosition =   3  '�t�ιw�]��
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
      TabIndex        =   13
      Text            =   " "
      Top             =   2160
      Width           =   732
   End
   Begin VB.CommandButton cmdDel 
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
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
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
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3360
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
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   855
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
      TabIndex        =   6
      Text            =   " "
      Top             =   1680
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
      TabIndex        =   5
      Text            =   " "
      Top             =   1200
      Width           =   1092
   End
   Begin VB.TextBox txtSeatNo 
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
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "�y        ��:"
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
      TabIndex        =   12
      Top             =   840
      Width           =   1095
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
      Top             =   2280
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
      Top             =   1800
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
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '��u�T�w
      Caption         =   " �� �� �� �Z ��� �@ �~"
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
Attribute VB_Name = "�H���ɪ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim student As studentrec  '�b(�@��)(�ŧi)��
Private Sub cmdOpen_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   ok = MsgBox("�p�G�w�g�}�ɹL,���ʧ@�|�M���ɮפ��e", vbOKCancel, "�M���ɮ�")
   If ok = vbOK Then
      student.seat_no = 0
      student.nam = ""
      student.chin = 0
      student.eng = 0
      For i = 1 To 100
         Put #1, i, student
      Next i
   End If
   Close #1
End Sub

Private Sub cmdAdd_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   student.seat_no = Val(txtSeatNo)
   If student.seat_no < 1 Or student.seat_no > 100 Then
      MsgBox "�y���W�X1~100���d��,�L�k�@�~"
   Else
      student.nam = txtName
      student.chin = Val(txtChin)
      student.eng = Val(txtEng)
      Put #1, student.seat_no, student
   End If
   Close #1
   txtSeatNo = "": txtName = "": txtChin = "": txtEng = ""
End Sub


Private Sub cmdQuery_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   num = Val(txtSeatNo)
   If num < 1 Or num > 100 Then
      MsgBox "�y���W�X1~100���d��,�L�k�@�~"
   Else
      Get #1, num, student
      If student.seat_no <> 0 Then
         txtName = student.nam
         txtChin = student.chin
         txtEng = student.eng
         MsgBox "�����,�äw���!"
      Else
         MsgBox "�ɮפ��S���n�䪺�y��:" + txtSeatNo
      End If
   End If
   Close #1
   txtSeatNo = "": txtName = "": txtChin = "": txtEng = ""
End Sub
Private Sub cmdDel_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   num = Val(txtSeatNo)
   If num < 1 Or num > 100 Then
      MsgBox "�y���W�X1~100���d��,�L�k�@�~"
   Else
      Get #1, num, student
      If student.seat_no <> 0 Then
         student.seat_no = 0
         student.nam = ""
         student.chin = 0
         student.eng = 0
         Put #1, num, student
         MsgBox "�����,�äw�R��!"
      Else
         MsgBox "�ɮפ��S���n�䪺�y��:" + txtSeatNo
      End If
   End If
   Close #1
   txtSeatNo = "": txtName = "": txtChin = "": txtEng = ""
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

