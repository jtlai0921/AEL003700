VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�p��q�O"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3690
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame Frame1 
      Caption         =   "�ιq����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton optBusiness 
         Caption         =   "��~��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optHome 
         Caption         =   "�a�x��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
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
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCal 
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
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtDegree 
      Alignment       =   1  '�a�k���
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Text            =   " "
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label LblFee 
      Alignment       =   1  '�a�k���
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
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   " �q�O:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�ιq�׼�:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   optHome.Value = True  '�w�]��ΤF�a�x�ιq
End Sub
Private Sub cmdCal_Click()
   Dim fee As Single, degree As Single
   degree = Val(txtDegree)
   If optHome.Value = True Then  '���p��a�x�ιq�N
      Call home(degree, fee)     '�I�s�Ƶ{��home
   Else                          '�_�h
      Call business(degree, fee) '�I�s�Ƶ{��Business
   End If
   LblFee = fee      '�N�q�O��J��ܵ��G����m
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

Private Sub home(d As Single, f As Single)
   Select Case d
      Case Is <= 100
         f = 2.4 * d
      Case Is <= 300
         f = 2.4 * 100 + 3.1 * (d - 100)
      Case Else
         f = 2.4 * 100 + 3.1 * 200 + 4.1 * (d - 300)
   End Select
End Sub
Private Sub business(d As Single, f As Single)
   If d <= 300 Then
      f = 5.9 * d
   Else
      f = 5.9 * 300 + 6.7 * (d - 300)
   End If
End Sub
