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
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除記錄"
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新記錄"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "最後一筆"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一筆"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "第一筆"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Data adoBasic 
      Caption         =   "基本資料"
      Connect         =   "Access"
      DatabaseName    =   "C:\db\student.mdb"
      DefaultCursorType=   0  '預設的資料指標
      DefaultType     =   2  '使用 ODBCDirect
      Exclusive       =   0   'False
      Height          =   405
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  '資料表(Table)
      RecordSource    =   "basic"
      Top             =   2460
      Width           =   3015
   End
   Begin VB.TextBox Text4 
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
      Left            =   960
      TabIndex        =   8
      Text            =   " "
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      Left            =   960
      TabIndex        =   6
      Text            =   " "
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox Text2 
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
      Left            =   3000
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   2
      Text            =   " "
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "電話"
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
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "地址"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "姓名"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "學號"
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
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "學 生 基 本 資 料"
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
Sub display()  '副程式,在「一般」「宣告」中輸入
   Text1 = adoBasic.Recordset("number")  '將資料表中四個資料欄
   Text2 = adoBasic.Recordset("name")    '的資料分別放到對應的
   Text3 = adoBasic.Recordset("address") '文字方塊，顯示出來
   Text4 = adoBasic.Recordset("tel")
End Sub


Private Sub cmdDelete_Click()
   adoBasic.Recordset.Delete   '刪除目前指標所在的記錄
   MsgBox "完成刪除工作", vbOKOnly, "刪除完成"
   Call cmdNext_Click
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

Private Sub cmdFirst_Click()
   adoBasic.Recordset.MoveFirst   '將指標移到第一筆記錄
   Call display
End Sub

Private Sub cmdLast_Click()
   adoBasic.Recordset.MoveLast  '將指標移到最後一筆記錄
   Call display
End Sub

Private Sub cmdNext_Click()
   adoBasic.Recordset.MoveNext   '將指標移到下一筆記錄
   If Not adoBasic.Recordset.EOF Then
      Call display
   Else
      MsgBox "已經在最後一筆紀錄，不能再往後移", vbOKOnly, "EOF"
   End If
End Sub

Private Sub cmdPrevious_Click()
   adoBasic.Recordset.MovePrevious   '將指標移到上一筆記錄
   If Not adoBasic.Recordset.BOF Then
      Call display
   Else
      MsgBox "已經在第一筆紀錄，不能再往前移", vbOKOnly, "BOF"
   End If
End Sub

Private Sub cmdUpdate_Click()
   adoBasic.Recordset.Edit    '編輯目前指標所在記錄的內容
   adoBasic.Recordset("number") = Text1
   adoBasic.Recordset("name") = Text2
   adoBasic.Recordset("address") = Text3
   adoBasic.Recordset("tel") = Text4
   adoBasic.Recordset.Update   '將整筆資料寫到目前指標所在記錄上
   MsgBox "完成更新工作", vbOKOnly, "更新完成"
End Sub
