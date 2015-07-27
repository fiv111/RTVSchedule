Attribute VB_Name = "RTVScheduleModule"
' 表の開始行列
Public Const kMarginRow = 8
Public Const kMarginCol = 10

' 開始日, 終了日のセル
Public Const kStartDay = "$E$1"
Public Const kFinDay = "$H$1"

' タスク番号, 担当の位置
Public Const kTask = "$B$1"
Public Const kStaff = "$D$1"

' 項目の範囲
Public Const kFirstItem = "$B$1"
Public Const kLastItem = "$H$1"

' 印刷用
' 印刷用範囲の開始セル
Public Const kPrintStartCell = "$B$2"
' 印刷用マージン
Public Const kPrintMargin = 28

' エラーメッセージ
Private Const kErrorMsg = "ERROR UUURRYYYY!!!!"
Private Const kSavedMsg = "以下の場所に保存しました。"

' 書き出し用の拡張子
' xlsx, pdf
Private Const kExtXLSX = ".xlsx"
Private Const kExtPDF = ".pdf"


' auto open
'Sub auto_open()
'  kColorGray = 1
'End Sub


' show autoDialog
Private Sub showDialog(msg, sec)
  Dim wsh As Object
  Set wsh = CreateObject("WScript.Shell")
  wsh.Popup msg, sec, "自動表示", vbInformation
  Set wsh = Nothing
End Sub


' Disable Screen Update and Calculation
Sub stopCalculate()
  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual
End Sub


' Enable Screen Update and Calculation
Sub startCalculate()
  Application.ScreenUpdating = True
  ActiveSheet.EnableCalculation = True
  Application.Calculation = xlCalculationAutomatic
End Sub


' getCell value.
Function getCell(cellId, Optional rStep = 0, Optional cStep = 0)
  r = Range(cellId).Row
  c = Range(cellId).Column
  getCell = Cells(r + rStep, c + cStep).Value
End Function


' Render the task number.
Function getTaskNumber(cellId)
  Set cell = Range(cellId).Offset(-1)

  If Len(cell) <= 0 Or Not IsNumeric(cell) Then
    cell = Range(cellId).Offset(-2)
  End If

  cell = cell + 1
  getTaskNumber = cell
  
  Set cell = Nothing
End Function


' Return an address of current cell.
' return currentCell
Function curt()
  curt = Evaluate("ADDRESS(ROW(), COLUMN())")
End Function


' Return a holiday list
Private Function getHoliday_()
  getHoliday_ = HolidaySheet.Range("A1", HolidaySheet.Cells(Rows.Count, 1).End(xlUp))
End Function


' Return the starting day.
Function getStartDay()
  h = getHoliday_()
  
  ' 一つ上の終了日
  Set e = Range(curt()).Offset(-1, 3)

  ' 前/後倒し期間
  Set p = Range(curt()).Offset(0, 2)

  ' 取得したいセル。空かどうかの比較用。
  Set c = Range(curt()).Offset(-1)

  ' 空の場合2つ上、ものがあるなら 1つ上のセルを取得
  If Len(c) <= 0 Then
    e = Range(curt()).Offset(-2, 3)
  Else
    e = Range(curt()).Offset(-1, 3)
  End If
  
  getStartDay = Application.WorksheetFunction.WorkDay(e, p, h)
  
  Set e = Nothing
  Set p = Nothing
  Set c = Nothing
End Function


' 終了日計算
Function getEndDay()
  Set s = Range(curt()).Offset(0, -3)
  Set p = Range(curt()).Offset(0, -2)
  h = getHoliday_()
  getEndDay = Application.WorksheetFunction.WorkDay(s, p - 1, h)
  
  Set s = Nothing
  Set p = Nothing
End Function


' セルの背景の塗りつぶす。
Private Sub drawCell_(r, col, color)
  Cells(r, col).Interior.color = color
End Sub


' ヘッダの灰色背景色
' drawHead_
Private Sub drawHead_(firstRow, lastRow)
  colorGray = RGB(166, 166, 166)
  Set taskRange = Range(kTask)
  taskRangeCol = taskRange.Column

  For i = firstRow To lastRow
    v = Cells(i, taskRangeCol).Value
    If Len(v) <= 0 Then
      r = Cells(i, taskRangeCol).Row
      Range(Cells(i, taskRangeCol), Cells(r, Columns.Count)).Interior.color = colorGray
    End If
  Next i
  
  Set taskRange = Nothing
End Sub


' スケジュールを設定する。
Sub drawSchedule()
  ' カレンダーの高さ, 全長
  lastRow = Cells(Rows.Count, kMarginRow).End(xlUp).Row
  lastCol = Cells(kMarginRow, Columns.Count).Column

  h = getHoliday_()

  ' スケジュールの最初の起点
  startCol = kMarginCol

  ' 背景をクリアする。
  Call clearBgColor_

  ' debug 用
  'lastRow = 9

  ' 各行を処理する。
  For r = kMarginRow To lastRow
    ' 各行の開始日付
    Set startDate = Cells(r, Range(kStartDay).Column)

    ' 各行の終了日付
    Set finDate = Cells(r, Range(kFinDay).Column)

    ' 開始日と終了日がエラーか日付以外の場合、エラーメッセージを出して終了させる。
    If IsError(startDate.Value) Or IsError(finDate.Value) Then
      Call showDialog(kErrorMsg, 1)
      Exit Sub
    Else
      ' ヘッダではない場合、処理を行う。
      If Len(startDate) > 0 Then

        ' 土日祝日込みの期間
        Period = finDate - startDate

        ' 工数所要期間
        workDayPeriod = startDate.Offset(0, 1)

        ' 各行の開始コラム
        startCol = kMarginCol + startDate - Range(kStartDay)

        ' 担当の背景色を取得する
        inChargeColor = getInChargeColor_(Cells(r, Range(kStaff).Column))

        ' 土日祝日含む期間
        For i = 0 To Period
          ' 毎日
          aDay = startDate + i

          'Debug.Print aDay
          ' 工数期間
          For j = 0 To workDayPeriod
            ' 休み以外の日
            wDay = CDate(WorksheetFunction.WorkDay(startDate, j, h))

            ' 対象日が wDay と一致した場合、背景を塗りつぶす。
            If aDay = wDay Then
              Call drawCell_(r, startCol + i, inChargeColor)
            End If

          Next j
        Next i

        ' B-H のセルも塗りつぶす。
        For i = Range(kFirstItem).Column To Range(kLastItem).Column
          Call drawCell_(r, i, inChargeColor)
        Next i

      End If
    End If
    
    Set startDate = Nothing
    Set finDate = Nothing
  Next r

  ' カレンダー描写
  Call drawCalendar_
End Sub


' draw calendar
Private Sub drawCalendar_()
  ' color
  colorBlue = RGB(146, 205, 220)
  colorRed = RGB(218, 150, 148)

  ' 開始日
  Set startDay = Range(kStartDay)

  ' 終了日
  Set finDay = Range(kFinDay)

  ' 合計日数
  totalDay = finDay - startDay

  ' 最後の行
  lastRow = Cells(Rows.Count, kMarginRow).End(xlUp).Row

  ' col 幅
  colWidth = 2.4

  ' 休み
  h = getHoliday_()

  ' I4:I6 の間
  num = kMarginRow - 2

  ' 描写開始 I4 から開始し、I6 まで終了
  For i = num - 2 To num
    ' 4, 5, 6
    r = Cells(i, kMarginCol).Row
    ' 9
    c = Cells(i, kMarginCol).Column

    ' I4-6 日付の部分をクリアする。
    Range(Cells(i, kMarginCol), Cells(r, Columns.Count)).ClearContents

    For j = 0 To totalDay
      ' 毎日
      d = startDay + j
      Set cell = Cells(r, c)

      ' 月の処理
      If i = 4 Then
        If Day(d) = 1 Or j = 0 Then
          cell.Value = d
          cell.ColumnWidth = colWidth
        End If
      ' 日と曜日の処理
      ElseIf i > 4 Then
        cell.Value = d
        cell.ColumnWidth = colWidth
      End If

      ' 最後の行だけやる。
      If i = num Then
        ' 土日祝日の背景塗りつぶし。
        For b = kMarginRow To lastRow

          ' 土
          If Weekday(d) = 7 Then
            Call drawCell_(b, c, colorBlue)
          ' 日
          ElseIf Weekday(d) = 1 Then
            Call drawCell_(b, c, colorRed)
          End If

          ' 祝日
          For Each k In h
            ' 日が祝日の場合、色をつける。
            If d = k Then
              Call drawCell_(b, c, colorRed)
            End If
          Next k

        Next b
      End If

      ' コラムを一個ずつずらす
      c = c + 1
      
      Set cell = Nothing
    Next j
  Next i

  ' 灰色背景
  Call drawHead_(kMarginRow - 1, lastRow)
  
  Set startDay = Nothing
  Set finDay = Nothing
End Sub


' 背景を削除する。
Private Sub clearBgColor_()
  ' カレンダーの開始セル
  startCell = Cells(kMarginRow - 1, kMarginCol).Address

  ' 不要な背景を削除する
  Range(startCell, Cells(Rows.Count, Columns.Count)).Interior.color = xlNone
End Sub


' 担当の色を調べる
Private Function getInChargeColor_(t)
  Set fc = t.FormatConditions

  ' 条件の合計数
  fLen = fc.Count
  cellColor = 0

  For i = 1 To fLen
    If fc(i).Formula1 = t.Formula Then
      cellColor = fc(i).Interior.color
    End If
  Next i

  getInChargeColor_ = cellColor
  
  Set fc = Nothing
End Function


Private Function schedulePrintArea_()
  ' 最後の列
  lastCol = Range(kFinDay) - Range(kStartDay) + kMarginCol

  ' 最後の行
  lastRow = Cells(Rows.Count, Range(kPrintStartCell).Row).End(xlUp).Row

  ' 印刷範囲
  Set scheduleRange = Range(Range(kPrintStartCell), Cells(lastRow, lastCol))

  ' 一回解除する。
  ActiveSheet.PageSetup.PrintArea = ""

  ' 印刷範囲を設定
  ActiveSheet.PageSetup.PrintArea = scheduleRange.Address

  schedulePrintArea_ = scheduleRange.Address
  
  Set scheduleRange = Nothing
End Function


' 保存先のファイル名を生成する
Private Function saveFileName_(ext)
  ' ファイル名
  fName = "schedule-" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss")

  ' 保存先はこのファイルと同じディレクトリにする。
  savePath = ThisWorkbook.Path & "\" & fName & ext

  saveFileName_ = savePath
End Function


' xlsx save
Sub saveXLSX()
  ' 保存先はこのファイルと同じディレクトリにする。
  savePath = saveFileName_(kExtXLSX)

  ' 描写停止
  Call stopCalculate

  ' 保存する前にスケジュールを一回更新する。
  Call drawSchedule

  ' スケジュール範囲を取得。
  Set area = Range(schedulePrintArea_())

  ' 範囲をコピー
  area.Copy

  ' 新しい workbook を作成
  Set wb = Workbooks.Add

  ' 新しい workbook の 有効なシートにペーストする。
  wb.ActiveSheet.PasteSpecial xlPasteAllUsingSourceTheme
  wb.ActiveSheet.PasteSpecial xlPasteColumnWidths
  wb.ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats

  'PasteSpecial
  ' 印刷の設定
  wb.ActiveSheet.PageSetup.PaperSize = xlPaperA3
  wb.ActiveSheet.PageSetup.Orientation = xlLandscape
  wb.ActiveSheet.PageSetup.TopMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.BottomMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.LeftMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.RightMargin = kPrintMargin

  ' xlsx に保存
  wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

  ' メッセージを出す。
  Call showDialog(kSavedMsg & savePath, 2)

  ' workbook を閉じる
  wb.Close

  ' コピペーモードを無効に
  Application.CutCopyMode = False

  ' 描写再開
  Call startCalculate
  
  Set wb = Nothing
  Set area = Nothing
End Sub


' saveAsPDF
Sub saveAsPDF()
  ' 保存先はこのファイルと同じディレクトリにする。
  savePath = saveFileName_(kExtPDF)

  ' 描写停止
  Call stopCalculate

  ' 保存する前にスケジュールを一回更新する。
  Call drawSchedule

  ' スケジュール範囲を取得。
  Set area = Range(schedulePrintArea_())

  ' 範囲をコピー
  'area.Copy

  ' 新しい workbook を作成
  'Set wb = Workbooks.Add

  ' 新しい workbook の 有効なシートにペーストする。
  'wb.ActiveSheet.PasteSpecial xlPasteAllUsingSourceTheme
  'wb.ActiveSheet.PasteSpecial xlPasteColumnWidths
  'wb.ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats

  Set s = ThisWorkbook.ActiveSheet

  'PasteSpecial
  ' 印刷の設定
  s.PageSetup.PaperSize = xlPaperA3
  s.PageSetup.Orientation = xlLandscape
  s.PageSetup.TopMargin = kPrintMargin
  s.PageSetup.BottomMargin = kPrintMargin
  s.PageSetup.LeftMargin = kPrintMargin
  s.PageSetup.RightMargin = kPrintMargin

  ' xlsx に保存
  s.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath

  ' メッセージを出す。
  Call showDialog(kSavedMsg & savePath, 2)

  ' コピペーモードを無効に
  Application.CutCopyMode = False

  ' 描写再開
  Call startCalculate
  
  Set s = Nothing
  Set area = Nothing
End Sub
