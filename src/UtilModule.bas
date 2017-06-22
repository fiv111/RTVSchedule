Attribute VB_Name = "UtilModule"
Option Explicit

Private Const kFontName = "Meiryo"
Private Const kFontSize = 10

Public Enum Digest
  SHA1
  SHA256
  SHA384
  SHA512
  HMACMD5
  HMACSHA1
  HMACSHA256
  HMACSHA384
  HMACSHA512
End Enum

' ---
' digest
' ---
' digestHash
Public Function digestHash(ByVal hashType As Integer, Optional ByVal secretKey As String = "", Optional ByVal str As String = "0", Optional ByVal lowercase = True) As String
  Dim digestObj, utf8Obj As Object
  Dim bytes() As Byte
  Dim hashList() As Byte
  Dim i, hash As Variant

  Set utf8Obj = CreateObject("System.Text.UTF8Encoding")
  Select Case hashType
  Case 0
    Set digestObj = CreateObject("System.Security.Cryptography.SHA1Managed")
  Case 1
    Set digestObj = CreateObject("System.Security.Cryptography.SHA256Managed")
  Case 2
    Set digestObj = CreateObject("System.Security.Cryptography.SHA384Managed")
  Case 3
    Set digestObj = CreateObject("System.Security.Cryptography.SHA512Managed")
  Case 4
    Set digestObj = CreateObject("System.Security.Cryptography.HMACMD5")
  Case 5
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA1")
  Case 6
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA256")
  Case 7
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA384")
  Case 8
    Set digestObj = CreateObject("System.Security.Cryptography.HMACSHA512")
  Case Else
    Set digestObj = CreateObject("System.Security.Cryptography.SHA1Managed")
  End Select
  
  bytes = utf8Obj.GetBytes_4(secretKey & str)
  hashList = digestObj.ComputeHash_2(bytes)
  
  For i = 1 To UBound(hashList) + 1
    hash = hash & Right("0" & Hex(AscB(MidB(hashList, i, 1))), 2)
  Next i
  
  If lowercase Then
    digestHash = LCase(hash)
  Else
    digestHash = hash
  End If
End Function



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
  Dim wb As Variant
  Do While Workbooks.Count >= 2
    For Each wb In Workbooks
      If wb.name <> ThisWorkbook.name Then
        Application.DisplayAlerts = False
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

' Get the information of currentsheet from the other worksheet.
Public Function getCurrentSheetInfo(ByVal sheetName As String, ByVal r As String, Optional ByVal num As Long = 0) As Variant
  Dim val As Variant

  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  With re
    .Pattern = "^[0-9]+$"
    .IgnoreCase = True
    .Global = True
  End With
  
  If re.test(ThisWorkbook.ActiveSheet.name) Then
    If num = 0 Then
      val = ThisWorkbook.Worksheets(sheetName).Range(r & CInt(ThisWorkbook.ActiveSheet.name) + 1).Value
    Else
      val = ThisWorkbook.Worksheets(sheetName).Range(r & num + 1).Value
    End If
  Else
    val = "-"
  End If

  If Len(val) <= 0 Then
    getCurrentSheetInfo = "-"
  Else
    getCurrentSheetInfo = val
  End If
End Function

' Create a attribute table.
Private Sub attachAttributeTable(Optional ByVal ac As Range = Nothing)
  ' selection
  Dim s As Range

  If ac Is Nothing Then
    Set s = ActiveCell
  Else
    Set s = ac
  End If

  If s.Count > 1 Then
    MsgBox ("More then one cell has been selected. Select only ONE cell. Try again.")
    Exit Sub
  End If

  ' head
  Dim docHead(5), attrHead(5), docVal(5) As String

  docHead(0) = "PageID"
  docHead(1) = "PageName"
  docHead(2) = "CreatedBy"
  docHead(3) = "UpdatedBy"
  docHead(4) = "CreatedAt"
  docHead(5) = "UpdatedAt"

  docVal(0) = "=getCurrentSheetInfo(" & Chr(34) & "Sitemap" & Chr(34) & ", " & Chr(34) & "A" & Chr(34) & ")"
  docVal(1) = "=getCurrentSheetInfo(" & Chr(34) & "Sitemap" & Chr(34) & ", " & Chr(34) & "B" & Chr(34) & ")"
  docVal(2) = "-"
  docVal(3) = "-"
  docVal(4) = Date
  docVal(5) = "=TODAY()"

  attrHead(0) = "ID"
  attrHead(1) = "Name"
  attrHead(2) = "Type"
  attrHead(3) = "Description"
  attrHead(4) = "URL"
  attrHead(5) = "RedirectURL"

  ' color
  Dim docInfoHeadColor, attrHeadColor, colorWhite, colorGrey As Long
  docInfoHeadColor = RGB(51, 102, 153)
  attrHeadColor = RGB(128, 128, 128)
  colorWhite = RGB(255, 255, 255)
  colorGrey = RGB(80, 80, 80)

  Call UtilModule.stopCalculate

  ' print document head
  Dim i As Variant
  Dim r, c As Long
  For i = 0 To UBound(docHead)
    r = s.Row
    c = s.Column + i

    With ActiveSheet.Cells(r, c)
      .Value = docHead(i)
      .Interior.color = docInfoHeadColor
      .Font.name = kFontName
      .Font.color = colorWhite
      .Font.Bold = True
      .Font.Size = kFontSize
      .Borders.color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With

    With ActiveSheet.Cells(r + 1, c)
      .Value = docVal(i)
      .Font.name = kFontName
      .Borders.color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With
  Next

  Dim j As Variant
  Dim attrSize As Integer
  attrSize = 10

  ' print attribute head
  For i = 0 To UBound(attrHead)
    r = s.Row + 2
    c = s.Column + i

    With ActiveSheet.Cells(r, c)
      .Value = attrHead(i)
      .Interior.color = attrHeadColor
      .Font.name = kFontName
      .Font.color = colorWhite
      .Font.Bold = True
      .Font.Size = kFontSize
      .Borders.color = colorGrey
      .Borders.Weight = xlThin
      .Borders.LineStyle = xlContinuous
    End With

    For j = 1 To attrSize
      If i = 0 Then
        ActiveSheet.Cells(r + j, c).Value = j
      Else
        ActiveSheet.Cells(r + j, c).Value = "-"
      End If

      With ActiveSheet.Cells(r + j, c)
        .Font.Size = kFontSize
        .Font.name = kFontName
        .Borders.color = colorGrey
        .Borders.Weight = xlThin
        .Borders.LineStyle = xlContinuous
      End With
    Next
  Next

  Call startCalculate
  Set s = Nothing
End Sub

' Create a attribute table for current worksheet.
Public Sub drawAttributeTable()
  attachAttributeTable ActiveCell
End Sub

' Create a attribute table for all worksheets.
Public Sub drawAttributeTableAll()
  Dim s As Range
  Set s = ActiveCell

  Dim shts As sheets
  Set shts = Worksheets

  Dim re As Object
  Set re = CreateObject("VBScript.RegExp")
  With re
    .Pattern = "^[0-9]+$"
    .IgnoreCase = True
    .Global = True
  End With
  
  Dim ws As Variant
  For Each ws In shts
    If re.test(ws.Name) Then
      ws.Activate
      attachAttributeTable s
    End If
  Next

  Set shts = Nothing
  Set s = Nothing
End Sub


' Rename worksheets to ordered number.
Public Sub mvWorksheets()
  Dim num, shtSize As Long
  num = 1

  Dim ss As sheets
  Set ss = ActiveWindow.SelectedSheets
  shtSize = ss.Count

  Dim i As Variant
  For i = num To shtSize
    ss.Item(i).name = Math.Rnd() * Now
  Next

  For i = num To shtSize
    ss.Item(i).name = i
  Next

  Set ss = Nothing
End Sub

' show hidden worksheets
Sub showWorksheets()
  Dim ws As sheets
  Set ws = ThisWorkbook.Worksheets

  Dim s As Variant
  For Each s In ws
    If Not s.Visible Then
      s.Visible = True
    End If
  Next

  Set ws = Nothing
End Sub



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
' last row, column
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
' Change All the shape object placement to be free-floating in those worksheets.
Public Sub shapesFreeFloating()
  Dim ws, s As Variant

  For Each ws In ThisWorkbook.Worksheets
    If ws.Shapes.Count > 0 Then
      For Each s In ws.Shapes
        If Not s.AutoShapeType = msoShapeMixed Then
          s.Placement = xlFreeFloating
        End If
      Next
    End If
  Next
End Sub

' Select all shape object from current worksheet.
Public Sub selectCurrentWorksheetShapes()
  If ActiveSheet.Shapes.Count > 0 Then
    ActiveSheet.Shapes.SelectAll
  End If
End Sub

' Attach a numbering label for selected shape.
Public Sub attachNumberingLabel()
  Dim sWidth, sHeight As Double
  Dim i As Variant
  Dim s, sp As Shape
  Dim labelName As String
  Dim bgColor, borderColor As Long

  ' Label Shape size. 1cm
  sWidth = 28.2
  sHeight = 28.2

  labelName = "VBAWFLabel"

  ' yellow bg
  ' black line
  bgColor = RGB(255, 255, 0)
  borderColor = RGB(0, 0, 0)

  If Not TypeName(Selection) = "Range" Then
    For i = 1 To Selection.ShapeRange.Count
      Set sp = Selection.ShapeRange.Item(i)
      Set s = ActiveSheet.Shapes.AddShape(msoShapeRectangle, sp.Left + sp.width - sWidth * 0.5, sp.Top - sHeight * 0.5, sWidth, sHeight)

      With s
        .name = labelName & i
        .Fill.ForeColor.RGB = bgColor
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 3
        .TextFrame.Characters.Text = i
        .TextFrame.Characters.Font.name = kFontName
        .TextFrame.Characters.Font.color = borderColor
        .TextFrame.Characters.Font.Size = kFontSize
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
      End With

      Set s = Nothing
      Set sp = Nothing
    Next
  End If
End Sub

' Attach numbering labels for all shapes on current worksheet.
Public Sub attachNumberingLabelAll()
  Call selectCurrentWorksheetShapes
  Call attachNumberingLabel
End Sub



' ---
' echo message
' ---
' show message
Public Sub pMsg(ByVal msg As String, Optional ByVal sec As Integer = 1)
  Dim o As Object
  Set o = CreateObject("WScript.Shell")
  o.Popup msg, sec, "Auto Display", vbInformation
  Set o = Nothing
End Sub


' ---
' html TAG
' ---
' tag
Public Function tag(ByVal tName As String, ByVal label As String, Optional ByVal styleObj As Object = Nothing) As String
  Dim doc As MSHTML.HTMLDocument
  Dim t As MSHTML.HTMLElementCollection

  Set doc = New MSHTML.HTMLDocument
  Set t = doc.createElement(tName)
  t.innerText = label

  If Not styleObj Is Nothing Then
    Dim s As Variant
    For Each s In styleObj
      t.style.setAttribute s, styleObj(s)
    Next
  End If

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
