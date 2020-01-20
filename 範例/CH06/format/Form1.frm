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
   AutoRedraw = True
   Print Format(12, "000")
   Print Format(-12, "000")
   Print Format(12, "####")
   Print Format(-12, "####")
   Print Format(12.3, "####.##")
   Print Format(-12.3, "####.##")
   Print Format(12.3, "####.00")
   Print Format(-12.3, "####.00")
   Print Format(0.423, "##.##%")
   Print Format(0.423, "##.00%")
   Print Format(1234567, "#####,###")
   Print Format(1234567.8, "#####,###.##")
   Print Format(12345.6, "$####,##0.00")
   Print Format(123, "++####")
   Print Format(123, "####+-+-")
   Print Format(123, "(####  --)")
   Print Format(1234567, "###/##/###")
   Print Format(Now, "d")
   Print Format(Now, "m-d")
   Print Format(Now, "yyyy")
   Print Format(Now, "m/d/yyyy")
   Print Format(Now, "h:m:s")
   Print Format(Now, "h:m:s AM/PM")
End Sub

