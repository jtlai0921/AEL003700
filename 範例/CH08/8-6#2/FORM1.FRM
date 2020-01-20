VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "切換"
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   972
   End
   Begin VB.PictureBox Picture2 
      Height          =   1452
      Left            =   360
      Picture         =   "FORM1.frx":0000
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   360
      Width           =   1212
   End
   Begin VB.PictureBox Picture1 
      Height          =   1452
      Left            =   2040
      ScaleHeight     =   1395
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Picture1.Visible = False Then
      Picture1.Picture = Picture2.Picture
      Picture1.Visible = True
      Picture2.Visible = False
   Else
      Picture2.Picture = Picture1.Picture
      Picture2.Visible = True
      Picture1.Visible = False
   End If
End Sub
