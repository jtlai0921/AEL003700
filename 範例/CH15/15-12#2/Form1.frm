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
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   Dim st As studentrec  '�ŧi�O���ܼ�
   Open "a:\student.dat" For Random As #1 Len = 18
   Print Tab(7); "*** �� �� �� �Z �� ***"
   Print "�y��"; Tab(7); "�m  �W"; Tab(17); "���"; _
                Tab(23); "�^��"; Tab(29); "����"
   For i = 1 To 100
      Get #1, i, st
      If st.seat_no <> 0 Then
         Print st.seat_no; Tab(7); st.nam; Tab(17); st.chin; _
         Tab(23); st.eng; Tab(29); (st.chin + st.eng) / 2
      End If
   Next i
   Close #1
End Sub

