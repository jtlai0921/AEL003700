VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4830
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "無效按鈕"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "有效按鈕"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "此按鈕會顯現"
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
