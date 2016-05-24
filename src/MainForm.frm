VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "ControlPanel"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---
' init
' ---
Private Sub UserForm_Initialize()
  ' textboxInit
  Call periodFrameTextboxInit
  Call fieldFrameInit
  
  With Me
    .Show vbModeless
  End With
End Sub


' ---
' PeriodFrame
' ---
' textbox init
Private Sub periodFrameTextboxInit()
  sd = MainModule.startDate
  ed = MainModule.endDate
  
  ' if startDate is empty(00:00:00) set date as today, else set date as input value.
  If sd = CDate("00:00:00") Then
    Set m = ThisWorkbook.Worksheets(MainModule.kMainSheetName).Range(MainModule.kDefaultStartDayAddr)
    If Len(m.Value) > 0 Then
      MainForm.PeriodFrame.StartDateText.Value = m.Value
    Else
      MainForm.PeriodFrame.StartDateText.Value = Date
    End If
    Set m = Nothing
  Else
    MainForm.PeriodFrame.StartDateText.Value = sd
  End If
  
  ' if endDate is empty(00:00:00) set date as today +31, else set date as input value.
  If ed = CDate("00:00:00") Then
    MainForm.PeriodFrame.EndDateText.Value = Date + 31
  Else
    MainForm.PeriodFrame.EndDateText.Value = ed
  End If
End Sub

Private Sub StartDateText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call loadCalendarForm
End Sub

Private Sub EndDateText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call loadCalendarForm
End Sub

' Render button on click handler.
' Set start/end date property, and refresh the screen (schedule).
Private Sub RenderButton_Click()
  Call MainModule.render
End Sub


' ---
' FieldFrame
' ---
Private Sub fieldFrameInit()
  MainForm.FieldFrame.BindingCheckBox.Value = False
  MainModule.binding = MainForm.FieldFrame.BindingCheckBox.Value
End Sub

Private Sub BindingCheckBox_Change()
  MainModule.binding = MainForm.FieldFrame.BindingCheckBox.Value
  
  If MainModule.binding = True Then
    MainForm.FieldFrame.UpdateButton.Visible = False
  Else
    MainForm.FieldFrame.UpdateButton.Visible = True
  End If
End Sub

Private Sub UpdateButton_Click()
  Call MainModule.updateTask
End Sub

' ---
' ExportFrame
' ---
Private Sub XlsxButton_Click()
  Call MainModule.saveAsXLSX
End Sub

Private Sub PdfButton_Click()
  Call MainModule.saveAsPDF
End Sub


' ---
' Private Method
' ---
' Call CalendarForm
Private Sub loadCalendarForm()
  Load CalendarForm
  CalendarForm.Show
End Sub
