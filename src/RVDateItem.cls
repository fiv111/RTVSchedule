VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RVDateItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' text color
Private color_ As Long

' background color
Private bgColor_ As Long

' font size
Private fontSize_ As Integer

' Value of an order number or day of week.
Private value_ As Variant

' Value of date
Private dateValue_ As Variant

' Set true if is holiday.
Private isHoliday_ As Boolean

Property Let color(ByVal val As Long)
  color_ = val
End Property

Property Let bgColor(ByVal val As Long)
  bgColor_ = val
End Property

Property Let fontSize(ByVal val As Integer)
  fontSize_ = val
End Property

Property Let value(ByVal val As Variant)
  value_ = val
End Property

Property Let dateValue(ByVal val As Variant)
  dateValue_ = val
End Property

Property Let isHoliday(ByVal val As Boolean)
  isHoliday_ = val
End Property

Property Get color() As Long
  color = color_
End Property

Property Get bgColor() As Long
  bgColor = bgColor_
End Property

Property Get fontSize() As Integer
  fontSize = fontSize_
End Property

Property Get value() As Variant
  value = value_
End Property

Property Get dateValue() As Variant
  dateValue = dateValue_
End Property

Property Get isHoliday() As Boolean
  isHoliday = isHoliday_
End Property


' ---
' init
' ---
Public Sub Class_Initialize()
  value = 0
  fontSize = 10
  color = RGB(0, 0, 0)
  bgColor = RGB(255, 255, 255)
  dateValue = 0
  isHoliday = False
End Sub

Public Sub init(ByVal iValue As Integer, ByVal iDateValue As Date, ByVal iIsHoliday As Boolean, ByVal iFontSize As Integer, ByVal iColor As Long, ByVal iBgColor As Long)
  value = iValue
  dateValue = iDateValue
  isHoliday = iIsHoliday
  fontSize = iFontSize
  color = iColor
  bgColor = iBgColor
End Sub
