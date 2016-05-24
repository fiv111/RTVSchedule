Attribute VB_Name = "MainModule"
' ---
' Const
' ---
' Cell address
Public Const kDefaultStartDayAddr = "$E$7"

' Number of start row and column
Private Const kStartRow = 6
Private Const kStartCol = 10

' Worksheet name
Public Const kMainSheetName = "main"
Private Const kHolidaySheetName = "holiday"
Private Const kStatusSheetName = "status"
Private Const kMemberSheetName = "member"

' Message
Private Const kErrorMsg = "ERROR UUURRYYYY!!!!"
Private Const kSavedMsg = "以下の場所に保存しました。"
Private Const kErrorData = "正しいデータではありません。確認してください。"

' File extension
Private Const kExtXLSX = ".xlsx"
Private Const kExtPDF = ".pdf"

' color
Private Const kSatColor = "146,205,220"
Private Const kSunColor = "218,150,148"
Private Const kTaskHeadColor = "166,166,166"
Private Const kHeadDefaultColor = "64,64,64"
Private Const kTodayColor = "0,204,153"

' Other
Private Const kDateCellWidth = 2.4
Private Const kYearMonthFormat = "yyyy/m"
Private Const kYYYYMDFormat = "yyyy/m/d"
Private Const kMonthFormat = "m"
Private Const kWeekdayFormat = "aaa"
Private Const kPrintMargin = 28

' ---
' Proerties
' ---
' The start/end date value from userform.
Dim startDate_ As Date
Dim endDate_ As Date

' Date binding flag
Dim binding_ As Boolean

' worksheets
Dim mainSheet_ As Worksheet
Dim holidaySheet_ As Worksheet
Dim memberSheet_ As Worksheet
Dim statusSheet_ As Worksheet

Private Property Let startDate(val As Date)
  startDate_ = val
End Property

Public Property Get startDate() As Date
  startDate = startDate_
End Property

Private Property Let endDate(val As Date)
  endDate_ = val
End Property

Public Property Get endDate() As Date
  endDate = endDate_
End Property

Public Property Let binding(val As Boolean)
  binding_ = val
End Property

Public Property Get binding() As Boolean
  binding = binding_
End Property

Private Property Let mainSheet(val As Worksheet)
  Set mainSheet_ = val
End Property

Private Property Get mainSheet() As Worksheet
  Set mainSheet = mainSheet_
End Property

Private Property Let holidaySheet(val As Worksheet)
  Set holidaySheet_ = val
End Property

Private Property Get holidaySheet() As Worksheet
  Set holidaySheet = holidaySheet_
End Property

Private Property Let memberSheet(val As Worksheet)
  Set memberSheet_ = val
End Property

Private Property Get memberSheet() As Worksheet
  Set memberSheet = memberSheet_
End Property

Private Property Let statusSheet(val As Worksheet)
  Set statusSheet_ = val
End Property

Private Property Get statusSheet() As Worksheet
  Set statusSheet = statusSheet_
End Property


' ---
' Public Method
' ---
' load MainForm
Public Sub loadMainForm()
Attribute loadMainForm.VB_ProcData.VB_Invoke_Func = "e\n14"
  Call worksheetInit
  Load MainForm
End Sub


' render schedule
Public Sub render()
  Call worksheetInit
  Call insertDefaultStartDate
  Call renderCalendar
End Sub


' update task cell
Public Sub updateTask()
  Call worksheetInit
  Call renderCalendar
End Sub


' startDay
Public Function getStartDate()
  ' 一つ上の終了日
  Set e = mainSheet.Range(UtilModule.curtAddr()).Offset(-1, 3)
  
  ' 前/後倒し期間
  Set p = Range(UtilModule.curtAddr()).Offset(0, 2)
  
  ' 取得したいセル。空かどうかの比較用。
  Set c = Range(UtilModule.curtAddr()).Offset(-1)

  ' 空の場合2つ上、ものがあるなら 1つ上のセルを取得
  If Len(c) <= 0 Then
    e = Range(UtilModule.curtAddr()).Offset(-2, 3)
  Else
    e = Range(UtilModule.curtAddr()).Offset(-1, 3)
  End If
  
  getStartDate = Application.WorksheetFunction.WorkDay(e, p, holidayRange())
  
  Set c = Nothing
  Set p = Nothing
  Set e = Nothing
End Function


' getEndDate
Public Function getEndDate()
  Set e = mainSheet.Range(UtilModule.curtAddr()).Offset(0, -3)
  Set p = mainSheet.Range(UtilModule.curtAddr()).Offset(0, -2)
  
  getEndDate = Application.WorksheetFunction.WorkDay(e, p - 1, holidayRange())
  Set p = Nothing
  Set e = Nothing
End Function


' ---
' Private Method
' ---
' auto open
Private Sub Auto_Open()
  Call UtilModule.stopCalculate
  
  Call worksheetInit
  Call loadMainForm
  
  mainSheet.Activate
  ActiveSheet.Calculate
  
  Call UtilModule.startCalculate
End Sub


' auto_close
Private Sub Auto_close()
  mainSheet = Nothing
  holidaySheet = Nothing
  memberSheet = Nothing
  statusSheet = Nothing
End Sub


' worksheetInit
Private Sub worksheetInit()
  Call setStartEndDate
  
  If mainSheet Is Nothing Then
    mainSheet = ThisWorkbook.Worksheets(kMainSheetName)
  End If
  
  If holidaySheet Is Nothing Then
    holidaySheet = ThisWorkbook.Worksheets(kHolidaySheetName)
  End If
  
  If memberSheet Is Nothing Then
    memberSheet = ThisWorkbook.Worksheets(kMemberSheetName)
  End If
  
  If statusSheet Is Nothing Then
    statusSheet = ThisWorkbook.Worksheets(kStatusSheetName)
  End If
  
  mainSheet.Activate
End Sub


' Set the start/end date value from MainForm.
Private Sub setStartEndDate()
  sd = MainForm.PeriodFrame.StartDateText.Value
  ed = MainForm.PeriodFrame.EndDateText.Value
  
  If IsDate(sd) And IsDate(ed) Then
    startDate_ = sd
    endDate_ = ed
  Else
    UtilModule.pMsg kErrorData, 1
  End If
End Sub


' set defaultStartDate
Private Sub insertDefaultStartDate()
  mainSheet.Range(kDefaultStartDayAddr).Value = startDate
End Sub


' renderCalendar
Private Sub renderCalendar()
  'Call UtilModule.stopCalculate
  
  ' clear background color
  Call clearBgColor
  Call drawDate
  Call drawHoliday(kSatColor, kSunColor)
  Call drawTaskHead(kTaskHeadColor)
  Call drawTask
  
  'ActiveSheet.Calculate
  'Call UtilModule.startCalculate
End Sub


' drawDate
Private Sub drawDate()
  ' position of start cell
  fRow = kStartRow - 3
  lRow = UtilModule.lastRow(mainSheet, fRow)
  fCol = kStartCol
  
  ' 3 lines of calendar value
  ' - month value
  ' - date value
  ' - week day value
  maxRow = fRow + 2
  
  ' totalDay
  totalDay = endDate - startDate
  
  Call UtilModule.stopCalculate
  
  For i = fRow To maxRow
    ' position of calendar cell
    cRow = mainSheet.Cells(i, fCol).Row
    cCol = mainSheet.Cells(i, fCol).Column
    
    ' clear all content before refresh those values.
    mainSheet.Range(Cells(i, fCol), Cells(cRow, Columns.Count)).ClearContents
    
    For j = 0 To totalDay
      ' everyday
      everyday = startDate + j
      Set c = mainSheet.Cells(cRow, cCol)
      
      ' default bg color
      Call drawCell(c, kHeadDefaultColor)
      
      ' Position of month.
      If i = fRow Then
        If Day(everyday) = 1 Or j = 0 Then
          ' print year when month is 1.
          If Format(everyday, kMonthFormat) = 1 Then
            c.Value = Format(everyday, kYearMonthFormat)
          Else
            c.Value = Format(everyday, kMonthFormat)
          End If
        End If
        
      ' Position of day.
      ElseIf i = fRow + 1 Then
        c.Value = Day(everyday)
        
      ' Position of week day.
      ElseIf i = maxRow Then
        c.Value = Format(everyday, kWeekdayFormat)
      End If
      
      ' is today
      If everyday = Date Then
        Call drawCell(c, kTodayColor)
      End If
      
      ' set column width
      c.ColumnWidth = kDateCellWidth
      
      ' next column
      cCol = cCol + 1
      
      Set c = Nothing
    Next
  Next
  
  ActiveSheet.Calculate
  Call UtilModule.startCalculate
End Sub


' drawHoliday
Private Sub drawHoliday(weekendColor, holidaydayColor)
  ' position of start cell
  fRow = kStartRow
  lRow = UtilModule.lastRow(mainSheet, fRow)
  fCol = kStartCol
  lCol = (endDate - startDate) + fCol
  everyday = startDate
  
  Call UtilModule.stopCalculate
  
  For i = fCol To lCol
    Set r = mainSheet.Range(mainSheet.Cells(fRow, i), mainSheet.Cells(lRow, i))
    
    ' Sat
    If Weekday(everyday) = 7 Then
      Call drawCell(r, weekendColor)
    End If
    
    ' sun
    If Weekday(everyday) = 1 Then
      Call drawCell(r, holidaydayColor)
    End If
    
    ' holiday
    Set hr = holidayRange().Find(What:=everyday, LookAt:=xlWhole, MatchCase:=True, MatchByte:=True)
    If Not hr Is Nothing Then
      Call drawCell(r, holidaydayColor)
    End If
    
    everyday = everyday + 1

    Set hr = Nothing
    Set r = Nothing
  Next
  
  ActiveSheet.Calculate
  Call UtilModule.startCalculate
End Sub


' Draw task head
Private Sub drawTaskHead(color)
  fRow = kStartRow
  fCol = 3
  lRow = UtilModule.lastRow(mainSheet, fRow)
  
  Call UtilModule.stopCalculate
  
  For i = fRow To lRow
    Set c = mainSheet.Cells(i, fCol)
    
    If Len(c.Offset(0, -1).Value) <= 0 _
    And Len(c.Offset(0, 1).Value) <= 0 _
    And c.Interior.color = UtilModule.getRGB(color) _
    Then
      Call drawCell(mainSheet.Range(mainSheet.Cells(i, fCol - 1), mainSheet.Cells(i, Columns.Count)), color)
    End If
    
    Set c = Nothing
  Next
  
  ActiveSheet.Calculate
  Call UtilModule.startCalculate
End Sub


' drawTask
Private Sub drawTask()
  fRow = kStartRow + 1
  fCol = kStartCol
  lRow = UtilModule.lastRow(mainSheet, fRow)
  lCol = (endDate - startDate) + fCol
  
  Call UtilModule.stopCalculate
  
  For i = fRow To lRow
    Set taskStartCell = mainSheet.Cells(i, fCol).Offset(0, -5)
    Set taskEndCell = taskStartCell.Offset(0, 3)
    Set wdayPeriodCell = taskStartCell.Offset(0, 1)
    
    If Not taskStartCell Is Nothing _
    Or Not taskEndCell Is Nothing _
    Or Not wdayPeriodCell Is Nothing Then
    
      If Not IsError(taskStartCell.Value) _
      And Not IsError(taskEndCell.Value) Then

        If Len(taskStartCell.Value) > 0 _
        And Len(taskEndCell.Value) > 0 Then
        
          period = taskEndCell.Value - taskStartCell.Value
          wdayPeriod = wdayPeriodCell.Value
          
          For j = 0 To period
            everyday = taskStartCell.Value + j
            
            For k = 0 To wdayPeriod
              wday = CDate(WorksheetFunction.WorkDay(taskStartCell.Value, k, holidayRange()))
              
              If everyday = wday Then
                ' task color
                taskColor = UtilModule.getInteriorColorByFormula(taskStartCell.Offset(0, -1))
                
                ' set taskColor in A:H
                Set taskRange = mainSheet.Range(taskStartCell.Offset(0, -3), taskStartCell.Offset(0, 3))
                taskRange.Interior.color = taskColor
                Set taskRange = Nothing
                
                ' status color
                statusColor = UtilModule.getInteriorColorByFormula(taskStartCell.Offset(0, 4))
                  
                ' set statusColor I:I
                taskStartCell.Offset(0, 4).Interior.color = statusColor
                
                ' 各タスクの開始カラム (デフォルトカラム + タスクの開始日 - スケジュール開始日)
                sCol = (fCol + taskStartCell.Value - startDate)
                If sCol > 0 Then
                  mainSheet.Cells(i, sCol + j).Interior.color = taskColor
                End If
              End If
              
            Next
          Next
          
        End If
      End If
    End If
    
    Set wdayPeriodCell = Nothing
    Set taskEndCell = Nothing
    Set taskStartCell = Nothing
  Next
  
  ActiveSheet.Calculate
  Call UtilModule.startCalculate
End Sub


' Draw cell background.
Private Sub drawCell(c, color)
  If LCase(TypeName(color)) = "string" Then
    c.Interior.color = UtilModule.getRGB(color)
  ElseIf LCase(TypeName(color)) = "long" Then
    c.Interior.color = color
  End If
End Sub


' Return holiday list
Private Function holidayRange()
  fRow = 2
  fCol = 2
  lRow = UtilModule.lastRow(holidaySheet, fRow)
  Set holidayRange = holidaySheet.Range(holidaySheet.Cells(fRow, fCol), holidaySheet.Cells(lRow, fCol))
End Function


' clear background color
Private Sub clearBgColor()
  Call UtilModule.stopCalculate
  Call drawCell(mainSheet.Range(mainSheet.Cells(kStartRow, kStartCol), mainSheet.Cells(Rows.Count, Columns.Count)), xlNone)
  ActiveSheet.Calculate
  Call UtilModule.startCalculate
End Sub


' ---
' Export file
' ---
Private Function getSaveFileName(ext)
  ' filename
  fName = "schedule-" & Format(Date, "yyyymmdd") & Format(time, "hhmmss")
  ' dest
  getSaveFileName = ThisWorkbook.Path & "\" & fName & ext
End Function


Private Function getPrintArea(sh)
  fRow = 1
  fCol = 1
  lRow = UtilModule.lastRow(sh, 2)
  lCol = UtilModule.lastCol(sh, 4)
  
  ActiveSheet.PageSetup.PrintArea = ""
  Set r = ActiveSheet.Range(Cells(1, 1), Cells(lRow, lCol))
  ActiveSheet.PageSetup.PrintArea = r.Address
  
  Set getPrintArea = r
  Set r = Nothing
End Function


' xlsx save
Public Sub saveAsXLSX()
  Call worksheetInit
  
  Set originBook = ThisWorkbook
  
  ' worksheet name list
  Set sheetlist = CreateObject("System.Collections.ArrayList")
  sheetlist.Add kStatusSheetName
  sheetlist.Add kMemberSheetName
  
  ' 保存先はこのファイルと同じディレクトリにする。
  saveDest = getSaveFileName(kExtXLSX)

  ' refresh the screen
  Call updateTask
  
  Call UtilModule.stopCalculate
  
  ' add new workbook
  On Error GoTo ExportError
  Dim wb As Workbook
  Set wb = Workbooks.Add
  wb.Worksheets(1).name = kMainSheetName
  
  ' mainsheet copy paste
  ' schedule area
  Set area = getPrintArea(mainSheet)
  area.Copy
  
  ' paste
  ' xlPasteValues
  ' xlPasteFormats
  ' xlPasteFormulas
  ' xlPasteAllMergingConditionalFormats
  wb.Activate
  wb.ActiveSheet.Select
  Selection.PasteSpecial xlPasteColumnWidths
  Selection.PasteSpecial xlPasteValidation
  Selection.PasteSpecial xlPasteAllUsingSourceTheme
  Selection.PasteSpecial xlPasteFormulasAndNumberFormats
  Selection.PasteSpecial xlPasteValuesAndNumberFormats
  wb.ActiveSheet.Select False
  Application.CutCopyMode = False
  
  ' print area setting
  wb.ActiveSheet.PageSetup.PaperSize = xlPaperA3
  wb.ActiveSheet.PageSetup.Orientation = xlLandscape
  wb.ActiveSheet.PageSetup.TopMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.BottomMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.LeftMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.RightMargin = kPrintMargin
  
  ' status/member sheet copy paste
  For Each s In sheetlist
    Dim wbSheet As Worksheet
    Set wbSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wbSheet.name = s
    
    ' copy
    originBook.Activate
    originBook.Worksheets(s).Cells.Copy
    
    ' paste
    wbSheet.Activate
    wbSheet.Paste
    Application.CutCopyMode = False
    
    wbSheet.Visible = xlSheetHidden
    
    Set wbSheet = Nothing
  Next
  
  ' activate the main sheet
  wb.Worksheets(kMainSheetName).Activate
  
  ' return to the origin workbook
  originBook.Activate
  
  ' save  and close the new workbook
  Application.DisplayAlerts = False
  wb.SaveAs Filename:=saveDest, FileFormat:=xlOpenXMLWorkbook
  Application.DisplayAlerts = True
  wb.Close
  
  ' done message
  Call UtilModule.pMsg(kSavedMsg & saveDest, 2)
  
  ' restart the auto calculation
  Call UtilModule.startCalculate
  
  ' free
  Set originBook = Nothing
  Set wb = Nothing
  Set area = Nothing
  Exit Sub
 
ExportError:
  Call UtilModule.pMsg("Worksheet add error." & kErrorMsg, 1)
  Call UtilModule.startCalculate
  Set originBook = Nothing
  Set wb = Nothing
  Set area = Nothing
End Sub


' ---
' export to pdf
' ---
Public Sub saveAsPDF()
  Call worksheetInit
  
  ' 保存先はこのファイルと同じディレクトリにする。
  saveDest = getSaveFileName(kExtPDF)
  
  ' refresh the screen
  Call updateTask
  
  Call UtilModule.stopCalculate
  
  'On Error GoTo ExportError
  Set s = ThisWorkbook.ActiveSheet
  
  ' print area
  Set area = getPrintArea(mainSheet)
  area.Select
  
  ' set printarea
  s.PageSetup.PaperSize = xlPaperA3
  s.PageSetup.Orientation = xlLandscape
  s.PageSetup.TopMargin = kPrintMargin
  s.PageSetup.BottomMargin = kPrintMargin
  s.PageSetup.LeftMargin = kPrintMargin
  s.PageSetup.RightMargin = kPrintMargin
  
  ' export 2 pdf
  Application.DisplayAlerts = False
  s.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveDest
  Application.DisplayAlerts = True
  
  ' done message
  Call UtilModule.pMsg(kSavedMsg & saveDest, 2)
    
  ' restart the auto calculation
  Call UtilModule.startCalculate
  
  ' free
  Set area = Nothing
  Set s = Nothing
  
  Exit Sub
 
ExportError:
  Call UtilModule.pMsg("ExportError." & kErrorMsg, 1)
  Call UtilModule.startCalculate
  Set area = Nothing
  Set s = Nothing
End Sub
