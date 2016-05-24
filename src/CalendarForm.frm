VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "Calendar"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3510
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---
' init
' ---
Private Sub UserForm_Initialize()
  CalendarPicker.Value = MainForm.PeriodFrame.ActiveControl.Value
End Sub


' ---
' Event
' ---
Private Sub CalendarPicker_Click()
  MainForm.PeriodFrame.ActiveControl.Value = CalendarPicker.Value
  Unload CalendarForm
End Sub
