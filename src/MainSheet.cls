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
