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
Option Explicit

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
  Dim sd As Date
  Dim ed As Date
  
  sd = CDate(Replace(Replace(ThisWorkbook.Names.item(MainModule.kStartDateName), "=", ""), Chr(34), ""))
  ed = CDate(Replace(Replace(ThisWorkbook.Names.item(MainModule.kEndDateName), "=", ""), Chr(34), ""))
  
  MainForm.PeriodFrame.StartDateText.value = sd
  MainForm.PeriodFrame.EndDateText.value = ed
End Sub

' ---
' Event
' ---
' Render button on click handler.
' Set start/end date property, and refresh the screen (schedule).
Private Sub RenderButton_Click()
  Call MainModule.render
End Sub


Private Sub StartDateText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Load RVCalendarForm
  RVCalendarForm.form = Me.PeriodFrame
End Sub

Private Sub EndDateText_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Load RVCalendarForm
  RVCalendarForm.form = Me.PeriodFrame
End Sub


' ---
' FieldFrame
' ---
Private Sub fieldFrameInit()
  MainForm.FieldFrame.BindingCheckBox.value = False
  MainModule.binding = MainForm.FieldFrame.BindingCheckBox.value
End Sub


Private Sub BindingCheckBox_Change()
  MainModule.binding = MainForm.FieldFrame.BindingCheckBox.value
  
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

