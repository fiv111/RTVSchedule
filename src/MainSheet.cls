VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
  If Intersect(Target, Range("$B:$H")) Is Nothing Then
    Exit Sub
  Else
  
    ' 自動計算停止
    Call RTVScheduleModule.stopCalculate
    ActiveSheet.Calculate
    Call RTVScheduleModule.startCalculate
      
    ' 開始日か終了日がいじられた場合は、カレンダーを更新する。
    If (Target.Column >= 4 And Target.Row >= 8) Or (Target.Address = RTVScheduleModule.kStartDay Or Target.Address = RTVScheduleModule.kFinDay) Then
      ' スケジュールを描写する。
      Call RTVScheduleModule.stopCalculate
      ActiveSheet.Calculate
      Call RTVScheduleModule.drawSchedule
    End If

    Call RTVScheduleModule.startCalculate
  End If
End Sub
