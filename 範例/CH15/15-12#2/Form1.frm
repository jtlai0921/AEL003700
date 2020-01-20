VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   Dim st As studentrec  '宣告記錄變數
   Open "a:\student.dat" For Random As #1 Len = 18
   Print Tab(7); "*** 學 生 成 績 表 ***"
   Print "座號"; Tab(7); "姓  名"; Tab(17); "國文"; _
                Tab(23); "英文"; Tab(29); "平均"
   For i = 1 To 100
      Get #1, i, st
      If st.seat_no <> 0 Then
         Print st.seat_no; Tab(7); st.nam; Tab(17); st.chin; _
         Tab(23); st.eng; Tab(29); (st.chin + st.eng) / 2
      End If
   Next i
   Close #1
End Sub

