VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�x�W�O���w"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
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
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   1290
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
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   1290
   End
   Begin VB.TextBox txtDays 
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
      TabIndex        =   10
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtDeposit 
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
      TabIndex        =   9
      Text            =   " "
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "��"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblTotalDeposit 
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
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
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "�����s���`�B"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblLastDeposit 
      Alignment       =   1  '�a�k���
      BorderStyle     =   1  '��u�T�w
      Caption         =   " "
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
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "�̫�@�����s���B"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "��"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "��"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "����s�ڤѼ�"
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
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "�Ĥ@�Ѧs���B"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
  Dim day_deposit As Double, total_deposit As Double  '��Ѧs�ڻP�ֿn�s��
  Dim days As Integer                       '����s�ڤѼ�
  deposit = Val(txtDeposit)
  days = Val(txtDays)
  total_deposit = deposit       '�Ĥ@�Ѫ��ֿn�s�ڵ����Ѧs��
  For i = 2 To days
    day_deposit = deposit * 2 ^ (i - 1) '�C�Ѧs�ڬ��Ĥ@�Ѧs�ڪ�2^(i-1)��
    total_deposit = total_deposit + day_deposit  '�ֿn�`�s�ڪ��B
  Next i
  Rem ��ܵ��G
  lblLastDeposit.Caption = Format(day_deposit, "#########,###")
  lblTotalDeposit.Caption = Format(total_deposit, "#########,###")
End Sub

Private Sub cmdEnd_Click()
   End
End Sub
