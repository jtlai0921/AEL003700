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
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   Dim a As String, b As String, c As String
   Dim boy As String, girl As String, join As String
   a = "Visual"
   b = "BASIC"
   c = a + " " + b + "�ܦn��"
   Print c
   boy = "ù�e��": girl = "���R��"
   join = boy & "�P" & girl
   Print join
   Print boy + "�P" + girl
End Sub

