VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   4590
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�R���O��"
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "��s�O��"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "�̫�@��"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "�W�@��"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "�U�@��"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "����"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "�Ĥ@��"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Data adoBasic 
      Caption         =   "�򥻸��"
      Connect         =   "Access"
      DatabaseName    =   "C:\db\student.mdb"
      DefaultCursorType=   0  '�w�]����ƫ���
      DefaultType     =   2  '�ϥ� ODBCDirect
      Exclusive       =   0   'False
      Height          =   405
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  '��ƪ�(Table)
      RecordSource    =   "basic"
      Top             =   2460
      Width           =   3015
   End
   Begin VB.TextBox Text4 
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
      Left            =   960
      TabIndex        =   8
      Text            =   " "
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      Left            =   960
      TabIndex        =   6
      Text            =   " "
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   855
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
      Left            =   960
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "�q��"
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
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "�a�}"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "�m�W"
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
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "�Ǹ�"
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
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "�� �� �� �� �� ��"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub display()  '�Ƶ{��,�b�u�@��v�u�ŧi�v����J
   Text1 = adoBasic.Recordset("number")  '�N��ƪ��|�Ӹ����
   Text2 = adoBasic.Recordset("name")    '����Ƥ��O��������
   Text3 = adoBasic.Recordset("address") '��r����A��ܥX��
   Text4 = adoBasic.Recordset("tel")
End Sub


Private Sub cmdDelete_Click()
   adoBasic.Recordset.Delete   '�R���ثe���ЩҦb���O��
   MsgBox "�����R���u�@", vbOKOnly, "�R������"
   Call cmdNext_Click
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

Private Sub cmdFirst_Click()
   adoBasic.Recordset.MoveFirst   '�N���в���Ĥ@���O��
   Call display
End Sub

Private Sub cmdLast_Click()
   adoBasic.Recordset.MoveLast  '�N���в���̫�@���O��
   Call display
End Sub

Private Sub cmdNext_Click()
   adoBasic.Recordset.MoveNext   '�N���в���U�@���O��
   If Not adoBasic.Recordset.EOF Then
      Call display
   Else
      MsgBox "�w�g�b�̫�@�������A����A���Ჾ", vbOKOnly, "EOF"
   End If
End Sub

Private Sub cmdPrevious_Click()
   adoBasic.Recordset.MovePrevious   '�N���в���W�@���O��
   If Not adoBasic.Recordset.BOF Then
      Call display
   Else
      MsgBox "�w�g�b�Ĥ@�������A����A���e��", vbOKOnly, "BOF"
   End If
End Sub

Private Sub cmdUpdate_Click()
   adoBasic.Recordset.Edit    '�s��ثe���ЩҦb�O�������e
   adoBasic.Recordset("number") = Text1
   adoBasic.Recordset("name") = Text2
   adoBasic.Recordset("address") = Text3
   adoBasic.Recordset("tel") = Text4
   adoBasic.Recordset.Update   '�N�㵧��Ƽg��ثe���ЩҦb�O���W
   MsgBox "������s�u�@", vbOKOnly, "��s����"
End Sub
