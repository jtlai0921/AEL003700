VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   FontSize = 14    '設定字型大小為14點
   Print Now       '輸出日期與時間
   Print Date       '輸出日期
   Print Time       '輸出時間
   Print            '空一列
   Print "民國";
   Print Val(Format(Date, "yyyy")) - 1911; _
         "年";
   Print Format(Date, "m"); "月";
   Print Format(Date, "d"); "日";
   Print Format(Time, "h"); "時";
   Print Format(Time, "n"); "分";
   Print Format(Time, "s"); "秒";
End Sub
