VERSION 5.00
Begin VB.Form 隨機檔的應用 
   Caption         =   "學生資料作業"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   3840
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox txtEng 
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
      Left            =   1440
      TabIndex        =   13
      Text            =   " "
      Top             =   2160
      Width           =   732
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除"
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
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
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
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增"
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
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
      Left            =   1440
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開檔"
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
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtChin 
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
      Left            =   1440
      TabIndex        =   6
      Text            =   " "
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtName 
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
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   1200
      Width           =   1092
   End
   Begin VB.TextBox txtSeatNo 
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
      Left            =   1440
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label5 
      Caption         =   "座        號:"
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
      TabIndex        =   12
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "英文成績:"
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
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "國文成績:"
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
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "姓        名:"
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
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      Caption         =   " 學 生 成 績 資料 作 業"
      BeginProperty Font 
         Name            =   "新細明體"
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
Attribute VB_Name = "隨機檔的應用"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim student As studentrec  '在(一般)(宣告)中
Private Sub cmdOpen_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   ok = MsgBox("如果已經開檔過,此動作會清除檔案內容", vbOKCancel, "清除檔案")
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
      MsgBox "座號超出1~100的範圍,無法作業"
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
      MsgBox "座號超出1~100的範圍,無法作業"
   Else
      Get #1, num, student
      If student.seat_no <> 0 Then
         txtName = student.nam
         txtChin = student.chin
         txtEng = student.eng
         MsgBox "有找到,並已顯示!"
      Else
         MsgBox "檔案中沒有要找的座號:" + txtSeatNo
      End If
   End If
   Close #1
   txtSeatNo = "": txtName = "": txtChin = "": txtEng = ""
End Sub
Private Sub cmdDel_Click()
   Open "a:\student.dat" For Random As #1 Len = 18
   num = Val(txtSeatNo)
   If num < 1 Or num > 100 Then
      MsgBox "座號超出1~100的範圍,無法作業"
   Else
      Get #1, num, student
      If student.seat_no <> 0 Then
         student.seat_no = 0
         student.nam = ""
         student.chin = 0
         student.eng = 0
         Put #1, num, student
         MsgBox "有找到,並已刪除!"
      Else
         MsgBox "檔案中沒有要找的座號:" + txtSeatNo
      End If
   End If
   Close #1
   txtSeatNo = "": txtName = "": txtChin = "": txtEng = ""
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

