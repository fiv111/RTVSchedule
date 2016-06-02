Attribute VB_Name = "UtilModule"
Option Explicit

' ---
' screen, calculate Update
' ---
' stopCalculate
Public Sub stopCalculate()
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
End Sub

' startCalculate
Public Sub startCalculate()
  Application.Calculation = xlCalculationAutomatic
  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating = True
End Sub


' ---
' workbooks
' ---
' close all workbooks
Public Sub closeAllBooks()
  Do While Workbooks.Count >= 2
    Dim wb As Variant
    For Each wb In Workbooks
      If wb.name <> ThisWorkbook.name Then
        Application.DisplayAlerts = Flase
        wb.Close saveChanges:=False
        Application.DisplayAlerts = True
      End If
    Next wb
  Loop
End Sub


' ---
' worksheet
' ---
' hasSheet
Public Function hasSheet(ByVal book As Workbook, ByVal name As String) As Boolean
  Dim s As Variant
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
' Cells
' ---
' current cell address
' Return an address of current cell.
Public Function curtAddr() As String
  curtAddr = Evaluate("ADDRESS(ROW(), COLUMN())")
End Function

' Get next number.
' Int cid
Public Function nextId(addr As String) As Double
  Dim c As Range
  Dim val As Variant

  Set c = ActiveSheet.Range(addr).Offset(-1)
  val = c.Value

  If Len(val) <= 0 Or Not IsNumeric(val) Then
    val = nextId(c.Address)
    GoTo done
  End If
  
  val = val + 1
  
done:
  nextId = val
  Set c = Nothing
End Function


' ---
' last row, col
' Worksheets o
' ---
' lastRow
Public Function lastRow(ByVal o As Worksheet, Optional ByVal first As Integer = 1) As Long
  lastRow = o.Cells(Rows.Count, first).End(xlUp).Row
End Function

' lastColoumn
Public Function lastCol(ByVal o As Worksheet, Optional ByVal first As Integer = 1) As Long
  lastCol = o.Cells(first, Columns.Count).End(xlToLeft).Column
End Function


' ---
' Sharpes
' ---
' set all sharpes position to be a fixed sharp.
Sub sharpFreeFloat()
  Dim ws As Variant
  
  For Each ws In ThisWorkbook.Worksheets
    
    If ws.Shapes.Count > 0 Then
      
      Dim s As Variant
      For Each s In ws.Shapes
        If Not s.AutoShapeType = msoShapeMixed Then
          s.Placement = xlFreeFloating
        End If
      Next
      
    End If
    
  Next
End Sub


' ---
' echo message
' ---
' show message
Public Sub pMsg(ByVal msg As String, Optional ByVal sec As Integer = 1)
  Dim o As Object
  Set o = CreateObject("WScript.Shell")
  o.Popup msg, sec, "自動表示", vbInformation
  Set o = Nothing
End Sub


' ---
' html TAG
' ---
' tag
Public Function tag(ByVal tName As String, ByVal label As String) As String
  Dim doc As MSHTML.HTMLDocument
  Dim t As MSHTML.HTMLElementCollection
  
  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement(tName)

  t.innerText = label
  tag = t.outerHTML

  Set t = Nothing
  Set doc = Nothing
End Function

' br
Public Function br() As String
  Dim doc As MSHTML.HTMLDocument
  Dim t As MSHTML.HTMLElementCollection
  
  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement("br")
  
  br = LCase(t.outerHTML)
  
  Set t = Nothing
  Set doc = Nothing
End Function


' ---
' glob
' ---
Public Sub glob(ByVal fPath As String, ByRef ary As Object)
  Dim fso As New Scripting.FileSystemObject
  
  Dim f As Variant
  For Each f In fso.GetFolder(fPath).files
    ary.Add f
  Next

  If fso.GetFolder(fPath).SubFolders.Count > 0 Then
    Dim d As Variant
    For Each d In fso.GetFolder(fPath).SubFolders
      ary.Add d
      glob d, ary
    Next
  End If
  
  Set fso = Nothing
End Sub


' ---
' array
' ---
' uniq
Public Function uniq(ByVal ary As Object) As Object
  Dim nAry As Object
  Set nAry = CreateObject("System.Collections.ArrayList")
  
  Dim v As Variant
  For Each v In ary
    If Not nAry.contains(v) Then
      nAry.Add v
    End If
  Next

  Set uniq = nAry
End Function


' ---
' color
' ---
' getRGB
' String c
'   ex: "255,255,255"
Public Function getRGB(Optional ByVal c As String = "0,0,0") As Long
  Dim myColor() As String
  myColor = Split(c, ",")
  getRGB = RGB(myColor(0), myColor(1), myColor(2))
End Function

' get interior color by Formula
Public Function getInteriorColorByFormula(ByVal c As Range) As Long
  Dim fc As FormatConditions
  Dim fLen As Long
  Dim cColor As Long
  Dim i As Integer
  
  Set fc = c.FormatConditions
  fLen = fc.Count
  cColor = 0

  For i = 1 To fLen
    If Evaluate(fc(i).Formula1) = c.Formula Then
      cColor = fc(i).Interior.color
      Exit For
    End If
  Next i

  getInteriorColorByFormula = cColor
  Set fc = Nothing
End Function


' ---
' Time
' ---
' String to (time)object
Public Function timeObject(Optional ByVal val As String = "00:00:00") As Object
  Dim o As Object
  Dim tmp() As String
  
  Set o = CreateObject("Scripting.Dictionary")
  
  tmp = Split(val, ":")
  
  o.Add "h", CDbl(tmp(0))
  o.Add "m", CDbl(tmp(1))
  o.Add "s", CDbl(tmp(2))
  
  Set timeObject = o
  
  Set o = Nothing
End Function

' Convert time value to second. (Long)
Public Function time2sec(ByVal h As Double, ByVal m As Double, ByVal s As Double) As Double
  Dim hourPerSec As Double
  Dim minPerSec As Double
  hourPerSec = 60 * 60
  minPerSec = 60
  time2sec = (h * hourPerSec) + (m * minPerSec) + s
End Function

' Convert string time value to second.
Public Function strtime2sec(ByVal strtime As String) As Double
  Dim t As Object
  Set t = timeObject(strtime)
  strtime2sec = time2sec(t("h"), t("m"), t("s"))
  Set t = Nothing
End Function

' Convert second to time string.
Public Function sec2time(ByVal s As Double) As String
  sec2time = Application.WorksheetFunction.Text(CDate(s / 86400#), "[h]:mm:ss")
End Function

' Return the first day in set month.
Public Function getFirstDayInMonth(ByVal today As Variant) As Variant
  getFirstDayInMonth = DateSerial(Year(today), Month(today), 1)
End Function

' Return the last day in set month.
Public Function getLastDayInMonth(ByVal today As Variant) As Variant
  getLastDayInMonth = DateSerial(Year(today), Month(today) + 1, 0)
End Function

' Return the workday in month
Public Function getWorkdayInMonth(ByVal firstDay As Variant, ByVal lastDay As Variant, ByVal holidayRange As Variant) As Double
  getWorkdayInMonth = Application.WorksheetFunction.NetworkDays(firstDay, lastDay, holidayRange)
End Function


' ---
' String
' ---
' reReplace
Public Function reReplace(ByVal val As String, ByVal rval As String, ByVal pat As String) As String
  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  
  With re
    .Pattern = pat
    .IgnoreCase = True
    .Global = True
  End With
  
  reReplace = re.Replace(val, rval)
  Set re = Nothing
End Function
