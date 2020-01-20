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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim st As studentrec
   Dim n As Integer
   Open "a:\student.dat" For Random As #2 Len = 18
   n = InputBox("請輸入記錄編號")
   Get 2, n, st
   Print "座號:"; st.seat_no
   Print "姓名:"; st.nam
   Print "國文成績:"; st.chin
   Print "英文成績:"; st.eng
   Print "平均成績:"; (st.chin + st.eng) / 2
   Close 2
End Sub

