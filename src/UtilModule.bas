Attribute VB_Name = "UtilModule"
' ---
' screen, calculate Update
' ---

' 自動更新停止
Public Sub stopCalculate()
  Application.ScreenUpdating = False
  'ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
End Sub


' 自動更新有効
Public Sub startCalculate()
  Application.ScreenUpdating = True
  'ActiveSheet.EnableCalculation = True
  Application.Calculation = xlCalculationAutomatic
End Sub


' ---
' book
' ---
' ほかのブックを開いている場合すべてを閉じる処理。
Public Sub closeAllBooks()
  Do While Workbooks.Count >= 2
    For Each wb In Workbooks
      If wb.name <> ThisWorkbook.name Then
        'Debug.Print wb.Name
        Application.DisplayAlerts = Flase
        wb.Close saveChanges:=False
        Application.DisplayAlerts = True
      End If
    Next wb
  Loop
End Sub


' ---
' last row, col
' ---
' lastRow
Public Function lastRow(o, Optional first As Integer = 1)
  lastRow = o.Cells(Rows.Count, first).End(xlUp).row
End Function


' lastCol
Public Function lastCol(o, Optional first As Integer = 1)
  lastCol = o.Cells(first, Columns.Count).End(xlToLeft).Column
End Function


' ---
' echo message
' ---
' show message
Public Sub pMsg(msg, sec)
  Dim o As Object
  Set o = CreateObject("WScript.Shell")
  o.Popup msg, sec, "自動表示", vbInformation
  Set o = Nothing
End Sub


' ---
' html TAG
' ---
' tag
Public Function tag(tName As String, str As String)
  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement(tName)
  t.innerText = str
  tag = t.outerHTML
  Set t = Nothing
  Set doc = Nothing
End Function


' br
Public Function br()
  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement("br")
  br = t.outerHTML
  Set t = Nothing
  Set doc = Nothing
End Function


' ---
' glob
' ---
Public Sub glob(fPath, ary)
  Dim fso As New Scripting.FileSystemObject
  
  For Each f In fso.GetFolder(fPath).files
    ary.add f
  Next

  If fso.GetFolder(fPath).SubFolders.Count > 0 Then
    For Each d In fso.GetFolder(fPath).SubFolders
      ary.add d
      glob d, ary
    Next
  End If

  Set fso = Nothing
End Sub


' ---
' worksheet
' ---
' hasSheet
Public Function hasSheet(book, ByVal name As String)
  For Each s In book.Worksheets
    If s.name = name Then
      hasSheet = True
      GoTo fin
    Else
      hasSheet = False
    End If
  Next
fin:
End Function


' ---
' array
' ---
' uniq
Function uniq(ary) As Object
  Set nAry = CreateObject("System.Collections.ArrayList")
  
  For Each v In ary
    If Not nAry.contains(v) Then
      nAry.add v
    End If
  Next
  
  Set uniq = nAry
End Function


'---
' color
' ---
' getRGB
Function getRGB(c)
  myColor = Split(c, ",")
  getRGB = RGB(myColor(0), myColor(1), myColor(2))
End Function
