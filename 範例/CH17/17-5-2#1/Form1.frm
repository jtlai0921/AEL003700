VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4410
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      DataField       =   "tel"
      DataSource      =   "Data1"
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
      Text            =   "Text4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "address"
      DataSource      =   "Data1"
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
      Text            =   "Text3"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Data1"
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
      Text            =   "Text2"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      DataField       =   "number"
      DataSource      =   "Data1"
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
      Text            =   "Text1"
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
