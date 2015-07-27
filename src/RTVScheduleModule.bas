Attribute VB_Name = "RTVScheduleModule"
' �\�̊J�n�s��
Public Const kMarginRow = 8
Public Const kMarginCol = 10

' �J�n��, �I�����̃Z��
Public Const kStartDay = "$E$1"
Public Const kFinDay = "$H$1"

' �^�X�N�ԍ�, �S���̈ʒu
Public Const kTask = "$B$1"
Public Const kStaff = "$D$1"

' ���ڂ͈̔�
Public Const kFirstItem = "$B$1"
Public Const kLastItem = "$H$1"

' ����p
' ����p�͈͂̊J�n�Z��
Public Const kPrintStartCell = "$B$2"
' ����p�}�[�W��
Public Const kPrintMargin = 28

' �G���[���b�Z�[�W
Private Const kErrorMsg = "ERROR UUURRYYYY!!!!"
Private Const kSavedMsg = "�ȉ��̏ꏊ�ɕۑ����܂����B"

' �����o���p�̊g���q
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
  wsh.Popup msg, sec, "�����\��", vbInformation
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
  
  ' ���̏I����
  Set e = Range(curt()).Offset(-1, 3)

  ' �O/��|������
  Set p = Range(curt()).Offset(0, 2)

  ' �擾�������Z���B�󂩂ǂ����̔�r�p�B
  Set c = Range(curt()).Offset(-1)

  ' ��̏ꍇ2��A���̂�����Ȃ� 1��̃Z�����擾
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


' �I�����v�Z
Function getEndDay()
  Set s = Range(curt()).Offset(0, -3)
  Set p = Range(curt()).Offset(0, -2)
  h = getHoliday_()
  getEndDay = Application.WorksheetFunction.WorkDay(s, p - 1, h)
  
  Set s = Nothing
  Set p = Nothing
End Function


' �Z���̔w�i�̓h��Ԃ��B
Private Sub drawCell_(r, col, color)
  Cells(r, col).Interior.color = color
End Sub


' �w�b�_�̊D�F�w�i�F
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


' �X�P�W���[����ݒ肷��B
Sub drawSchedule()
  ' �J�����_�[�̍���, �S��
  lastRow = Cells(Rows.Count, kMarginRow).End(xlUp).Row
  lastCol = Cells(kMarginRow, Columns.Count).Column

  h = getHoliday_()

  ' �X�P�W���[���̍ŏ��̋N�_
  startCol = kMarginCol

  ' �w�i���N���A����B
  Call clearBgColor_

  ' debug �p
  'lastRow = 9

  ' �e�s����������B
  For r = kMarginRow To lastRow
    ' �e�s�̊J�n���t
    Set startDate = Cells(r, Range(kStartDay).Column)

    ' �e�s�̏I�����t
    Set finDate = Cells(r, Range(kFinDay).Column)

    ' �J�n���ƏI�������G���[�����t�ȊO�̏ꍇ�A�G���[���b�Z�[�W���o���ďI��������B
    If IsError(startDate.Value) Or IsError(finDate.Value) Then
      Call showDialog(kErrorMsg, 1)
      Exit Sub
    Else
      ' �w�b�_�ł͂Ȃ��ꍇ�A�������s���B
      If Len(startDate) > 0 Then

        ' �y���j�����݂̊���
        Period = finDate - startDate

        ' �H�����v����
        workDayPeriod = startDate.Offset(0, 1)

        ' �e�s�̊J�n�R����
        startCol = kMarginCol + startDate - Range(kStartDay)

        ' �S���̔w�i�F���擾����
        inChargeColor = getInChargeColor_(Cells(r, Range(kStaff).Column))

        ' �y���j���܂ފ���
        For i = 0 To Period
          ' ����
          aDay = startDate + i

          'Debug.Print aDay
          ' �H������
          For j = 0 To workDayPeriod
            ' �x�݈ȊO�̓�
            wDay = CDate(WorksheetFunction.WorkDay(startDate, j, h))

            ' �Ώۓ��� wDay �ƈ�v�����ꍇ�A�w�i��h��Ԃ��B
            If aDay = wDay Then
              Call drawCell_(r, startCol + i, inChargeColor)
            End If

          Next j
        Next i

        ' B-H �̃Z�����h��Ԃ��B
        For i = Range(kFirstItem).Column To Range(kLastItem).Column
          Call drawCell_(r, i, inChargeColor)
        Next i

      End If
    End If
    
    Set startDate = Nothing
    Set finDate = Nothing
  Next r

  ' �J�����_�[�`��
  Call drawCalendar_
End Sub


' draw calendar
Private Sub drawCalendar_()
  ' color
  colorBlue = RGB(146, 205, 220)
  colorRed = RGB(218, 150, 148)

  ' �J�n��
  Set startDay = Range(kStartDay)

  ' �I����
  Set finDay = Range(kFinDay)

  ' ���v����
  totalDay = finDay - startDay

  ' �Ō�̍s
  lastRow = Cells(Rows.Count, kMarginRow).End(xlUp).Row

  ' col ��
  colWidth = 2.4

  ' �x��
  h = getHoliday_()

  ' I4:I6 �̊�
  num = kMarginRow - 2

  ' �`�ʊJ�n I4 ����J�n���AI6 �܂ŏI��
  For i = num - 2 To num
    ' 4, 5, 6
    r = Cells(i, kMarginCol).Row
    ' 9
    c = Cells(i, kMarginCol).Column

    ' I4-6 ���t�̕������N���A����B
    Range(Cells(i, kMarginCol), Cells(r, Columns.Count)).ClearContents

    For j = 0 To totalDay
      ' ����
      d = startDay + j
      Set cell = Cells(r, c)

      ' ���̏���
      If i = 4 Then
        If Day(d) = 1 Or j = 0 Then
          cell.Value = d
          cell.ColumnWidth = colWidth
        End If
      ' ���Ɨj���̏���
      ElseIf i > 4 Then
        cell.Value = d
        cell.ColumnWidth = colWidth
      End If

      ' �Ō�̍s�������B
      If i = num Then
        ' �y���j���̔w�i�h��Ԃ��B
        For b = kMarginRow To lastRow

          ' �y
          If Weekday(d) = 7 Then
            Call drawCell_(b, c, colorBlue)
          ' ��
          ElseIf Weekday(d) = 1 Then
            Call drawCell_(b, c, colorRed)
          End If

          ' �j��
          For Each k In h
            ' �����j���̏ꍇ�A�F������B
            If d = k Then
              Call drawCell_(b, c, colorRed)
            End If
          Next k

        Next b
      End If

      ' �R������������炷
      c = c + 1
      
      Set cell = Nothing
    Next j
  Next i

  ' �D�F�w�i
  Call drawHead_(kMarginRow - 1, lastRow)
  
  Set startDay = Nothing
  Set finDay = Nothing
End Sub


' �w�i���폜����B
Private Sub clearBgColor_()
  ' �J�����_�[�̊J�n�Z��
  startCell = Cells(kMarginRow - 1, kMarginCol).Address

  ' �s�v�Ȕw�i���폜����
  Range(startCell, Cells(Rows.Count, Columns.Count)).Interior.color = xlNone
End Sub


' �S���̐F�𒲂ׂ�
Private Function getInChargeColor_(t)
  Set fc = t.FormatConditions

  ' �����̍��v��
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
  ' �Ō�̗�
  lastCol = Range(kFinDay) - Range(kStartDay) + kMarginCol

  ' �Ō�̍s
  lastRow = Cells(Rows.Count, Range(kPrintStartCell).Row).End(xlUp).Row

  ' ����͈�
  Set scheduleRange = Range(Range(kPrintStartCell), Cells(lastRow, lastCol))

  ' ����������B
  ActiveSheet.PageSetup.PrintArea = ""

  ' ����͈͂�ݒ�
  ActiveSheet.PageSetup.PrintArea = scheduleRange.Address

  schedulePrintArea_ = scheduleRange.Address
  
  Set scheduleRange = Nothing
End Function


' �ۑ���̃t�@�C�����𐶐�����
Private Function saveFileName_(ext)
  ' �t�@�C����
  fName = "schedule-" & Format(Date, "yyyymmdd") & Format(Time, "hhmmss")

  ' �ۑ���͂��̃t�@�C���Ɠ����f�B���N�g���ɂ���B
  savePath = ThisWorkbook.Path & "\" & fName & ext

  saveFileName_ = savePath
End Function


' xlsx save
Sub saveXLSX()
  ' �ۑ���͂��̃t�@�C���Ɠ����f�B���N�g���ɂ���B
  savePath = saveFileName_(kExtXLSX)

  ' �`�ʒ�~
  Call stopCalculate

  ' �ۑ�����O�ɃX�P�W���[�������X�V����B
  Call drawSchedule

  ' �X�P�W���[���͈͂��擾�B
  Set area = Range(schedulePrintArea_())

  ' �͈͂��R�s�[
  area.Copy

  ' �V���� workbook ���쐬
  Set wb = Workbooks.Add

  ' �V���� workbook �� �L���ȃV�[�g�Ƀy�[�X�g����B
  wb.ActiveSheet.PasteSpecial xlPasteAllUsingSourceTheme
  wb.ActiveSheet.PasteSpecial xlPasteColumnWidths
  wb.ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats

  'PasteSpecial
  ' ����̐ݒ�
  wb.ActiveSheet.PageSetup.PaperSize = xlPaperA3
  wb.ActiveSheet.PageSetup.Orientation = xlLandscape
  wb.ActiveSheet.PageSetup.TopMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.BottomMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.LeftMargin = kPrintMargin
  wb.ActiveSheet.PageSetup.RightMargin = kPrintMargin

  ' xlsx �ɕۑ�
  wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook

  ' ���b�Z�[�W���o���B
  Call showDialog(kSavedMsg & savePath, 2)

  ' workbook �����
  wb.Close

  ' �R�s�y�[���[�h�𖳌���
  Application.CutCopyMode = False

  ' �`�ʍĊJ
  Call startCalculate
  
  Set wb = Nothing
  Set area = Nothing
End Sub


' saveAsPDF
Sub saveAsPDF()
  ' �ۑ���͂��̃t�@�C���Ɠ����f�B���N�g���ɂ���B
  savePath = saveFileName_(kExtPDF)

  ' �`�ʒ�~
  Call stopCalculate

  ' �ۑ�����O�ɃX�P�W���[�������X�V����B
  Call drawSchedule

  ' �X�P�W���[���͈͂��擾�B
  Set area = Range(schedulePrintArea_())

  ' �͈͂��R�s�[
  'area.Copy

  ' �V���� workbook ���쐬
  'Set wb = Workbooks.Add

  ' �V���� workbook �� �L���ȃV�[�g�Ƀy�[�X�g����B
  'wb.ActiveSheet.PasteSpecial xlPasteAllUsingSourceTheme
  'wb.ActiveSheet.PasteSpecial xlPasteColumnWidths
  'wb.ActiveSheet.PasteSpecial xlPasteValuesAndNumberFormats

  Set s = ThisWorkbook.ActiveSheet

  'PasteSpecial
  ' ����̐ݒ�
  s.PageSetup.PaperSize = xlPaperA3
  s.PageSetup.Orientation = xlLandscape
  s.PageSetup.TopMargin = kPrintMargin
  s.PageSetup.BottomMargin = kPrintMargin
  s.PageSetup.LeftMargin = kPrintMargin
  s.PageSetup.RightMargin = kPrintMargin

  ' xlsx �ɕۑ�
  s.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath

  ' ���b�Z�[�W���o���B
  Call showDialog(kSavedMsg & savePath, 2)

  ' �R�s�y�[���[�h�𖳌���
  Application.CutCopyMode = False

  ' �`�ʍĊJ
  Call startCalculate
  
  Set s = Nothing
  Set area = Nothing
End Sub
