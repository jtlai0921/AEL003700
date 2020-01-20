VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "調色盤"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "調色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   855
   End
   Begin VB.PictureBox picRGB 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   2760
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtBlue 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "0 "
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtGreen 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtRed 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "0"
      Top             =   705
      Width           =   495
   End
   Begin VB.HScrollBar hsbBlue 
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   255
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.HScrollBar hsbGreen 
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   255
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.HScrollBar hsbRed 
      Height          =   255
      LargeChange     =   10
      Left            =   1320
      Max             =   255
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "255"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "數值設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "捲軸調整"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "調出的顏色:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "藍   色"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "綠   色"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "紅   色"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "三原色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdColor_Click()
   picRGB.BackColor = RGB(Val(txtRed), _
          Val(txtGreen), Val(txtBlue))    '利用此函數調色再顯示
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

Private Sub hsbBlue_Change()
  txtBlue = hsbBlue.Value '設定文字方塊的內容為捲動鈕的位置值
End Sub

Private Sub hsbGreen_Change()
  txtGreen = hsbGreen.Value '設定文字方塊的內容為捲動鈕的位置值
End Sub

Private Sub hsbRed_Change()
  txtRed = hsbRed.Value  '設定文字方塊的內容為捲動鈕的位置值
End Sub

