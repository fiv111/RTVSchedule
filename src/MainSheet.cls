Private Sub Worksheet_Change(ByVal Target As Range)
  If Intersect(Target, Range("$A:$I")) Is Nothing Then
    Exit Sub
  End If
  
  If MainModule.binding = True Then
    If (Target.Column >= 5 And Target.Column <= 8) And (Target.Row > 6) Then
      Call MainModule.updateTask
    End If
  Else
    Call UtilModule.stopCalculate
    ActiveSheet.Calculate
    Call UtilModule.startCalculate
  End If
End Sub
