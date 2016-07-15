VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RVCalendarForm 
   Caption         =   "RVCalendar"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   OleObjectBlob   =   "RVCalendarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "RVCalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cal_ As RVCalendar
Private form_ As MSForms.UserForm

' ---
' Property
' ---
Property Let cal(ByVal val As RVCalendar)
  Set cal_ = val
End Property

Property Let form(ByVal val As MSForms.UserForm)
  Set form_ = val
End Property

Property Get cal() As RVCalendar
  Set cal = cal_
End Property

Property Get form() As MSForms.UserForm
  Set form = form_
End Property

Property Get formatDate() As String
  formatDate = Format(cal.calDate, "mmm yyyy")
End Property


' ---
' init
' ---
Private Sub UserForm_Initialize()
  Call init
  
  With Me
    .Show vbModeless
  End With
End Sub


Private Sub init()
  cal = New RVCalendar
  cal.init Date
  cal.outputType = "form"
  
  Me.currentDateLabel.Caption = formatDate
  
  Call renderHead
  Call refresh
End Sub


Private Sub renderHead()
  Dim hList As Object
  Dim cHead As Object
  Dim size As Integer
  Dim labelName As String
  
  labelName = "hLabel"
  Set hList = labelList(labelName, cal.weekSize)
  Set cHead = cal.calHead()
  size = hList.Count - 1
  
  Dim i As Variant
  For i = 0 To size
    hList.item(i).Caption = cHead.item(i).value
    hList.item(i).BackColor = cHead.item(i).bgColor
    hList.item(i).ForeColor = cHead.item(i).color
  Next
  
  Set cHead = Nothing
  Set hList = Nothing
End Sub


Private Sub refresh()
  Dim dLabelList As Object
  Dim labelName As String
  Dim mSize As Integer
  Dim dItem As Variant
  Dim cLabel As Variant

  labelName = "dLabel"
  Set dLabelList = labelList(labelName, cal.monthSize)
  
  mSize = cal.monthlyList.Count - 1
  
  Dim i As Variant
  For i = 0 To mSize
    Set dItem = cal.monthlyList.item(i)
    Set cLabel = dLabelList.item(i)
    
    If dItem.value = 0 Then
      cLabel.Caption = ""
    Else
      cLabel.Caption = dItem.value
    End If
    cLabel.BackColor = dItem.bgColor
    cLabel.Font.size = dItem.fontSize

    Set cLabel = Nothing
    Set dItem = Nothing
  Next
  
  Set dLabelList = Nothing
End Sub


' Return a labelList find by name
Private Function labelList(ByVal labelName As String, ByVal size As Integer) As Object
  Dim ary As Object
  Set ary = CreateObject("System.Collections.ArrayList")
  
  Dim i As Variant
  For i = 0 To size - 1
    ary.Add RVCalendarForm.Controls.item(labelName & i)
  Next
  
  Set labelList = ary
  Set ary = Nothing
End Function


' ---
' Event
' ---
Private Sub currentDateLabel_Click()
  cal.calDate = Date
  currentDateLabel.Caption = formatDate
  Call refresh
End Sub


Private Sub prevLabel_Click()
  cal.calDate = cal.prevMonth
  currentDateLabel.Caption = formatDate
  Call refresh
End Sub

Private Sub nextLabel_Click()
  cal.calDate = cal.nextMonth
  currentDateLabel.Caption = formatDate
  Call refresh
End Sub

Private Sub dLabel0_Click()
  Call labelHander(dLabel0, 0)
End Sub

Private Sub dLabel1_Click()
  Call labelHander(dLabel1, 1)
End Sub

Private Sub dLabel2_Click()
  Call labelHander(dLabel2, 2)
End Sub

Private Sub dLabel3_Click()
  Call labelHander(dLabel3, 3)
End Sub

Private Sub dLabel4_Click()
  Call labelHander(dLabel4, 4)
End Sub

Private Sub dLabel5_Click()
  Call labelHander(dLabel5, 5)
End Sub

Private Sub dLabel6_Click()
  Call labelHander(dLabel6, 6)
End Sub

Private Sub dLabel7_Click()
  Call labelHander(dLabel7, 7)
End Sub

Private Sub dLabel8_Click()
  Call labelHander(dLabel8, 8)
End Sub

Private Sub dLabel9_Click()
  Call labelHander(dLabel9, 9)
End Sub

Private Sub dLabel10_Click()
  Call labelHander(dLabel10, 10)
End Sub

Private Sub dLabel11_Click()
  Call labelHander(dLabel11, 11)
End Sub

Private Sub dLabel12_Click()
  Call labelHander(dLabel12, 12)
End Sub

Private Sub dLabel13_Click()
  Call labelHander(dLabel13, 13)
End Sub

Private Sub dLabel14_Click()
  Call labelHander(dLabel14, 14)
End Sub

Private Sub dLabel15_Click()
  Call labelHander(dLabel15, 15)
End Sub

Private Sub dLabel16_Click()
  Call labelHander(dLabel16, 16)
End Sub

Private Sub dLabel17_Click()
  Call labelHander(dLabel17, 17)
End Sub

Private Sub dLabel18_Click()
  Call labelHander(dLabel18, 18)
End Sub

Private Sub dLabel19_Click()
  Call labelHander(dLabel19, 19)
End Sub

Private Sub dLabel20_Click()
  Call labelHander(dLabel20, 20)
End Sub

Private Sub dLabel21_Click()
  Call labelHander(dLabel21, 21)
End Sub

Private Sub dLabel22_Click()
  Call labelHander(dLabel22, 22)
End Sub

Private Sub dLabel23_Click()
  Call labelHander(dLabel23, 23)
End Sub

Private Sub dLabel24_Click()
  Call labelHander(dLabel24, 24)
End Sub

Private Sub dLabel25_Click()
  Call labelHander(dLabel25, 25)
End Sub

Private Sub dLabel26_Click()
  Call labelHander(dLabel26, 26)
End Sub

Private Sub dLabel27_Click()
  Call labelHander(dLabel27, 27)
End Sub

Private Sub dLabel28_Click()
  Call labelHander(dLabel28, 28)
End Sub

Private Sub dLabel29_Click()
  Call labelHander(dLabel29, 29)
End Sub

Private Sub dLabel30_Click()
  Call labelHander(dLabel30, 30)
End Sub

Private Sub dLabel31_Click()
  Call labelHander(dLabel31, 31)
End Sub

Private Sub dLabel32_Click()
  Call labelHander(dLabel32, 32)
End Sub

Private Sub dLabel33_Click()
  Call labelHander(dLabel33, 33)
End Sub

Private Sub dLabel34_Click()
  Call labelHander(dLabel34, 34)
End Sub

Private Sub dLabel35_Click()
  Call labelHander(dLabel35, 35)
End Sub

Private Sub dLabel36_Click()
  Call labelHander(dLabel36, 36)
End Sub

Private Sub dLabel37_Click()
  Call labelHander(dLabel37, 37)
End Sub

Private Sub dLabel38_Click()
  Call labelHander(dLabel38, 38)
End Sub

Private Sub dLabel39_Click()
  Call labelHander(dLabel39, 39)
End Sub

Private Sub dLabel40_Click()
  Call labelHander(dLabel40, 40)
End Sub

Private Sub dLabel41_Click()
  Call labelHander(dLabel41, 41)
End Sub

Private Sub labelHander(ByVal o As MSForms.label, ByVal num As Integer)
  If Len(o.Caption) > 0 Then
    cal.selectedDate = cal.monthlyList.item(num).dateValue
    
    If cal.outputType = "book" Then
      ActiveCell.value = cal.selectedDate
    Else
      If Not form Is Nothing Then
        If Not form.ActiveControl Is Nothing Then
          form.ActiveControl.value = cal.selectedDate
        End If
      End If
    End If
    
    Unload Me
  End If
End Sub
