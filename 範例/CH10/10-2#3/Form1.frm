VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1140
      ItemData        =   "Form1.frx":0000
      Left            =   840
      List            =   "Form1.frx":0019
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Print List1.ListCount
   Print List1.List(4)
   Print List1.List(1)
   Print List1.Text
   Print List1.ListIndex
   Print List1.Selected(3)
   Print List1.Selected(4)
End Sub


