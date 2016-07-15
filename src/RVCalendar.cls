VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RVCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Const
Private Const kWeekSize As Integer = 7
' 6w * 7d
Private Const kMonthSize As Integer = 42

' Date, year, month, day
Private calDate_ As Date
Private yy_ As Integer
Private mm_ As Integer
Private dd_ As Integer

' Using in userform.
' Will set after a label click event.
Private selectedDate_ As Date

' color
' Saturday color
Private saturdayBgColor_ As Long
Private saturdayColor_ As Long

' Holiday color
Private offdayBgColor_ As Long
Private offdayColor_ As Long

' onday color
Private ondayBgColor_ As Long
Private ondayColor_ As Long

' head or something like that
Private fieldBgColor_ As Long
Private fieldColor_ As Long

' highlight color
Private highlightBgColor_ As Long
Private highlightColor_ As Long

' fontsize
Private fontSize_ As Integer

' lang
Private lang_ As String

' holidayArray
Private holidayList_ As Object

' outputType
' return a value for book | form
Private outputType_ As String

' ---
' Property
' ---
' Let
Property Let calDate(ByVal val As Date)
  calDate_ = val
  calYear = year(calDate)
  calMonth = month(calDate)
  calDay = day(calDate)
End Property

Property Let calYear(ByVal val As Integer)
  yy_ = val
End Property

Property Let calMonth(ByVal val As Integer)
  mm_ = val
End Property

Property Let calDay(ByVal val As Integer)
  dd_ = val
End Property

Property Let saturdayBgColor(ByVal val As Long)
  saturdayBgColor_ = val
End Property

Property Let saturdayColor(ByVal val As Long)
  saturdayColor_ = val
End Property

Property Let ondayBgColor(ByVal val As Long)
  ondayBgColor_ = val
End Property

Property Let ondayColor(ByVal val As Long)
  ondayColor_ = val
End Property

Property Let offdayBgColor(ByVal val As Long)
  offdayBgColor_ = val
End Property

Property Let offdayColor(ByVal val As Long)
  offdayColor_ = val
End Property

Property Let fieldBgColor(ByVal val As Long)
  fieldBgColor_ = val
End Property

Property Let highlightBgColor(ByVal val As Long)
  highlightBgColor_ = val
End Property

Property Let highlightColor(ByVal val As Long)
  highlightColor_ = val
End Property

Property Let fieldColor(ByVal val As Long)
  fieldColor_ = val
End Property

Property Let fontSize(ByVal val As Integer)
  fontSize_ = val
End Property

Property Let lang(ByVal val As String)
  lang_ = val
End Property

Property Let holidayList(ByVal val As Object)
  Set holidayList_ = val
End Property

Property Let outputType(ByVal val As String)
  outputType_ = val
End Property

Property Let selectedDate(ByVal val As Date)
  selectedDate_ = val
End Property

' Getter
Property Get calDate() As Date
  calDate = calDate_
End Property

Property Get calYear() As Integer
  calYear = yy_
End Property

Property Get calMonth() As Integer
  calMonth = mm_
End Property

Property Get calDay() As Integer
  calDay = dd_
End Property

Property Get calWday() As Integer
  calWday = Weekday(calDate)
End Property

Property Get prevYear() As Date
  prevYear = DateSerial(year(calDate) - 1, month(calDate), day(calDate))
End Property

Property Get nextYear() As Date
  nextYear = DateSerial(year(calDate) + 1, month(calDate), day(calDate))
End Property

Property Get prevMonth() As Date
  prevMonth = DateSerial(year(calDate), month(calDate) - 1, day(calDate))
End Property

Property Get nextMonth() As Date
  nextMonth = DateSerial(year(calDate), month(calDate) + 1, day(calDate))
End Property

Property Get prevDay() As Date
  prevDay = DateSerial(year(calDate), month(calDate), day(calDate) - 1)
End Property

Property Get nextDay() As Date
  nextDay = DateSerial(year(calDate), month(calDate), day(calDate) + 1)
End Property

Property Get saturdayBgColor() As Long
  saturdayBgColor = saturdayBgColor_
End Property

Property Get saturdayColor() As Long
  saturdayColor = saturdayColor_
End Property

Property Get ondayBgColor() As Long
  ondayBgColor = ondayBgColor_
End Property

Property Get ondayColor() As Long
  ondayColor = ondayColor_
End Property

Property Get offdayBgColor() As Long
  offdayBgColor = offdayBgColor_
End Property

Property Get offdayColor() As Long
  offdayColor = offdayColor_
End Property

Property Get fieldBgColor() As Long
  fieldBgColor = fieldBgColor_
End Property

Property Get fieldColor() As Long
  fieldColor = fieldColor_
End Property

Property Get highlightBgColor() As Long
  highlightBgColor = highlightBgColor_
End Property

Property Get highlightColor() As Long
  highlightColor = highlightColor_
End Property

Property Get fontSize() As Integer
  fontSize = fontSize_
End Property

Property Get lang() As String
  lang = lang_
End Property

Property Get holidayList() As Object
  Set holidayList = holidayList_
End Property

Property Get weekSize() As Integer
  weekSize = kWeekSize
End Property

Property Get monthSize() As Integer
  monthSize = kMonthSize
End Property

Property Get outputType() As String
  outputType = outputType_
End Property

Property Get selectedDate() As Date
  selectedDate = selectedDate_
End Property

' ---
' init
' ---
Public Sub Class_Initialize()
  saturdayBgColor  = RGB(0, 125, 204)
  saturdayColor    = RGB(255, 255, 255)
  offdayBgColor    = RGB(204, 0, 0)
  offdayColor      = RGB(255, 255, 255)
  ondayBgColor     = RGB(100, 100, 100)
  ondayColor       = RGB(255, 255, 255)
  fieldBgColor     = RGB(50, 50, 50)
  fieldColor       = RGB(255, 255, 255)
  highlightBgColor = RGB(204, 190, 0)
  highlightColor   = RGB(255, 255, 255)
  fontSize         = 10
  outputType       = "book"
  lang             = "ja"

  ' this is a japanese holiday list from 2013 to 2032
  holidayList = CreateObject("System.Collections.ArrayList")
  holidayList.Add CDate("2013/1/1")
  holidayList.Add CDate("2013/1/14")
  holidayList.Add CDate("2013/2/11")
  holidayList.Add CDate("2013/3/20")
  holidayList.Add CDate("2013/4/29")
  holidayList.Add CDate("2013/5/3")
  holidayList.Add CDate("2013/5/4")
  holidayList.Add CDate("2013/5/5")
  holidayList.Add CDate("2013/5/6")
  holidayList.Add CDate("2013/7/15")
  holidayList.Add CDate("2013/9/16")
  holidayList.Add CDate("2013/9/23")
  holidayList.Add CDate("2013/10/14")
  holidayList.Add CDate("2013/11/3")
  holidayList.Add CDate("2013/11/4")
  holidayList.Add CDate("2013/11/23")
  holidayList.Add CDate("2013/12/23")
  holidayList.Add CDate("2014/1/1")
  holidayList.Add CDate("2014/1/13")
  holidayList.Add CDate("2014/2/11")
  holidayList.Add CDate("2014/3/21")
  holidayList.Add CDate("2014/4/29")
  holidayList.Add CDate("2014/5/3")
  holidayList.Add CDate("2014/5/4")
  holidayList.Add CDate("2014/5/5")
  holidayList.Add CDate("2014/5/6")
  holidayList.Add CDate("2014/7/21")
  holidayList.Add CDate("2014/9/15")
  holidayList.Add CDate("2014/9/23")
  holidayList.Add CDate("2014/10/13")
  holidayList.Add CDate("2014/11/3")
  holidayList.Add CDate("2014/11/23")
  holidayList.Add CDate("2014/11/24")
  holidayList.Add CDate("2014/12/23")
  holidayList.Add CDate("2015/1/1")
  holidayList.Add CDate("2015/1/12")
  holidayList.Add CDate("2015/2/11")
  holidayList.Add CDate("2015/3/21")
  holidayList.Add CDate("2015/4/29")
  holidayList.Add CDate("2015/5/3")
  holidayList.Add CDate("2015/5/4")
  holidayList.Add CDate("2015/5/5")
  holidayList.Add CDate("2015/5/6")
  holidayList.Add CDate("2015/7/20")
  holidayList.Add CDate("2015/9/21")
  holidayList.Add CDate("2015/9/22")
  holidayList.Add CDate("2015/9/23")
  holidayList.Add CDate("2015/10/12")
  holidayList.Add CDate("2015/11/3")
  holidayList.Add CDate("2015/11/23")
  holidayList.Add CDate("2015/12/23")
  holidayList.Add CDate("2016/1/1")
  holidayList.Add CDate("2016/1/11")
  holidayList.Add CDate("2016/2/11")
  holidayList.Add CDate("2016/3/20")
  holidayList.Add CDate("2016/3/21")
  holidayList.Add CDate("2016/4/29")
  holidayList.Add CDate("2016/5/3")
  holidayList.Add CDate("2016/5/4")
  holidayList.Add CDate("2016/5/5")
  holidayList.Add CDate("2016/7/18")
  holidayList.Add CDate("2016/8/11")
  holidayList.Add CDate("2016/9/19")
  holidayList.Add CDate("2016/9/22")
  holidayList.Add CDate("2016/10/10")
  holidayList.Add CDate("2016/11/3")
  holidayList.Add CDate("2016/11/23")
  holidayList.Add CDate("2016/12/23")
  holidayList.Add CDate("2017/1/1")
  holidayList.Add CDate("2017/1/2")
  holidayList.Add CDate("2017/1/9")
  holidayList.Add CDate("2017/2/11")
  holidayList.Add CDate("2017/3/20")
  holidayList.Add CDate("2017/4/29")
  holidayList.Add CDate("2017/5/3")
  holidayList.Add CDate("2017/5/4")
  holidayList.Add CDate("2017/5/5")
  holidayList.Add CDate("2017/7/17")
  holidayList.Add CDate("2017/8/11")
  holidayList.Add CDate("2017/9/18")
  holidayList.Add CDate("2017/9/23")
  holidayList.Add CDate("2017/10/9")
  holidayList.Add CDate("2017/11/3")
  holidayList.Add CDate("2017/11/23")
  holidayList.Add CDate("2017/12/23")
  holidayList.Add CDate("2018/1/1")
  holidayList.Add CDate("2018/1/8")
  holidayList.Add CDate("2018/2/11")
  holidayList.Add CDate("2018/2/12")
  holidayList.Add CDate("2018/3/21")
  holidayList.Add CDate("2018/4/29")
  holidayList.Add CDate("2018/4/30")
  holidayList.Add CDate("2018/5/3")
  holidayList.Add CDate("2018/5/4")
  holidayList.Add CDate("2018/5/5")
  holidayList.Add CDate("2018/7/16")
  holidayList.Add CDate("2018/8/11")
  holidayList.Add CDate("2018/9/17")
  holidayList.Add CDate("2018/9/23")
  holidayList.Add CDate("2018/9/24")
  holidayList.Add CDate("2018/10/8")
  holidayList.Add CDate("2018/11/3")
  holidayList.Add CDate("2018/11/23")
  holidayList.Add CDate("2018/12/23")
  holidayList.Add CDate("2018/12/24")
  holidayList.Add CDate("2019/1/1")
  holidayList.Add CDate("2019/1/14")
  holidayList.Add CDate("2019/2/11")
  holidayList.Add CDate("2019/3/21")
  holidayList.Add CDate("2019/4/29")
  holidayList.Add CDate("2019/5/3")
  holidayList.Add CDate("2019/5/4")
  holidayList.Add CDate("2019/5/5")
  holidayList.Add CDate("2019/5/6")
  holidayList.Add CDate("2019/7/15")
  holidayList.Add CDate("2019/8/11")
  holidayList.Add CDate("2019/8/12")
  holidayList.Add CDate("2019/9/16")
  holidayList.Add CDate("2019/9/23")
  holidayList.Add CDate("2019/10/14")
  holidayList.Add CDate("2019/11/3")
  holidayList.Add CDate("2019/11/4")
  holidayList.Add CDate("2019/11/23")
  holidayList.Add CDate("2019/12/23")
  holidayList.Add CDate("2020/1/1")
  holidayList.Add CDate("2020/1/13")
  holidayList.Add CDate("2020/2/11")
  holidayList.Add CDate("2020/3/20")
  holidayList.Add CDate("2020/4/29")
  holidayList.Add CDate("2020/5/3")
  holidayList.Add CDate("2020/5/4")
  holidayList.Add CDate("2020/5/5")
  holidayList.Add CDate("2020/5/6")
  holidayList.Add CDate("2020/7/20")
  holidayList.Add CDate("2020/8/11")
  holidayList.Add CDate("2020/9/21")
  holidayList.Add CDate("2020/9/22")
  holidayList.Add CDate("2020/10/12")
  holidayList.Add CDate("2020/11/3")
  holidayList.Add CDate("2020/11/23")
  holidayList.Add CDate("2020/12/23")
  holidayList.Add CDate("2021/1/1")
  holidayList.Add CDate("2021/1/11")
  holidayList.Add CDate("2021/2/11")
  holidayList.Add CDate("2021/3/20")
  holidayList.Add CDate("2021/4/29")
  holidayList.Add CDate("2021/5/3")
  holidayList.Add CDate("2021/5/4")
  holidayList.Add CDate("2021/5/5")
  holidayList.Add CDate("2021/7/19")
  holidayList.Add CDate("2021/8/11")
  holidayList.Add CDate("2021/9/20")
  holidayList.Add CDate("2021/9/23")
  holidayList.Add CDate("2021/10/11")
  holidayList.Add CDate("2021/11/3")
  holidayList.Add CDate("2021/11/23")
  holidayList.Add CDate("2021/12/23")
  holidayList.Add CDate("2022/1/1")
  holidayList.Add CDate("2022/1/10")
  holidayList.Add CDate("2022/2/11")
  holidayList.Add CDate("2022/3/21")
  holidayList.Add CDate("2022/4/29")
  holidayList.Add CDate("2022/5/3")
  holidayList.Add CDate("2022/5/4")
  holidayList.Add CDate("2022/5/5")
  holidayList.Add CDate("2022/7/18")
  holidayList.Add CDate("2022/8/11")
  holidayList.Add CDate("2022/9/19")
  holidayList.Add CDate("2022/9/23")
  holidayList.Add CDate("2022/10/10")
  holidayList.Add CDate("2022/11/3")
  holidayList.Add CDate("2022/11/23")
  holidayList.Add CDate("2022/12/23")
  holidayList.Add CDate("2023/1/1")
  holidayList.Add CDate("2023/1/2")
  holidayList.Add CDate("2023/1/9")
  holidayList.Add CDate("2023/2/11")
  holidayList.Add CDate("2023/3/21")
  holidayList.Add CDate("2023/4/29")
  holidayList.Add CDate("2023/5/3")
  holidayList.Add CDate("2023/5/4")
  holidayList.Add CDate("2023/5/5")
  holidayList.Add CDate("2023/7/17")
  holidayList.Add CDate("2023/8/11")
  holidayList.Add CDate("2023/9/18")
  holidayList.Add CDate("2023/9/23")
  holidayList.Add CDate("2023/10/9")
  holidayList.Add CDate("2023/11/3")
  holidayList.Add CDate("2023/11/23")
  holidayList.Add CDate("2023/12/23")
  holidayList.Add CDate("2024/1/1")
  holidayList.Add CDate("2024/1/8")
  holidayList.Add CDate("2024/2/11")
  holidayList.Add CDate("2024/2/12")
  holidayList.Add CDate("2024/3/20")
  holidayList.Add CDate("2024/4/29")
  holidayList.Add CDate("2024/5/3")
  holidayList.Add CDate("2024/5/4")
  holidayList.Add CDate("2024/5/5")
  holidayList.Add CDate("2024/5/6")
  holidayList.Add CDate("2024/7/15")
  holidayList.Add CDate("2024/8/11")
  holidayList.Add CDate("2024/8/12")
  holidayList.Add CDate("2024/9/16")
  holidayList.Add CDate("2024/9/22")
  holidayList.Add CDate("2024/9/23")
  holidayList.Add CDate("2024/10/14")
  holidayList.Add CDate("2024/11/3")
  holidayList.Add CDate("2024/11/4")
  holidayList.Add CDate("2024/11/23")
  holidayList.Add CDate("2024/12/23")
  holidayList.Add CDate("2025/1/1")
  holidayList.Add CDate("2025/1/13")
  holidayList.Add CDate("2025/2/11")
  holidayList.Add CDate("2025/3/20")
  holidayList.Add CDate("2025/4/29")
  holidayList.Add CDate("2025/5/3")
  holidayList.Add CDate("2025/5/4")
  holidayList.Add CDate("2025/5/5")
  holidayList.Add CDate("2025/5/6")
  holidayList.Add CDate("2025/7/21")
  holidayList.Add CDate("2025/8/11")
  holidayList.Add CDate("2025/9/15")
  holidayList.Add CDate("2025/9/23")
  holidayList.Add CDate("2025/10/13")
  holidayList.Add CDate("2025/11/3")
  holidayList.Add CDate("2025/11/23")
  holidayList.Add CDate("2025/11/24")
  holidayList.Add CDate("2025/12/23")
  holidayList.Add CDate("2026/1/1")
  holidayList.Add CDate("2026/1/12")
  holidayList.Add CDate("2026/2/11")
  holidayList.Add CDate("2026/3/20")
  holidayList.Add CDate("2026/4/29")
  holidayList.Add CDate("2026/5/3")
  holidayList.Add CDate("2026/5/4")
  holidayList.Add CDate("2026/5/5")
  holidayList.Add CDate("2026/5/6")
  holidayList.Add CDate("2026/7/20")
  holidayList.Add CDate("2026/8/11")
  holidayList.Add CDate("2026/9/21")
  holidayList.Add CDate("2026/9/22")
  holidayList.Add CDate("2026/9/23")
  holidayList.Add CDate("2026/10/12")
  holidayList.Add CDate("2026/11/3")
  holidayList.Add CDate("2026/11/23")
  holidayList.Add CDate("2026/12/23")
  holidayList.Add CDate("2027/1/1")
  holidayList.Add CDate("2027/1/11")
  holidayList.Add CDate("2027/2/11")
  holidayList.Add CDate("2027/3/21")
  holidayList.Add CDate("2027/3/22")
  holidayList.Add CDate("2027/4/29")
  holidayList.Add CDate("2027/5/3")
  holidayList.Add CDate("2027/5/4")
  holidayList.Add CDate("2027/5/5")
  holidayList.Add CDate("2027/7/19")
  holidayList.Add CDate("2027/8/11")
  holidayList.Add CDate("2027/9/20")
  holidayList.Add CDate("2027/9/23")
  holidayList.Add CDate("2027/10/11")
  holidayList.Add CDate("2027/11/3")
  holidayList.Add CDate("2027/11/23")
  holidayList.Add CDate("2027/12/23")
  holidayList.Add CDate("2028/1/1")
  holidayList.Add CDate("2028/1/10")
  holidayList.Add CDate("2028/2/11")
  holidayList.Add CDate("2028/3/20")
  holidayList.Add CDate("2028/4/29")
  holidayList.Add CDate("2028/5/3")
  holidayList.Add CDate("2028/5/4")
  holidayList.Add CDate("2028/5/5")
  holidayList.Add CDate("2028/7/17")
  holidayList.Add CDate("2028/8/11")
  holidayList.Add CDate("2028/9/18")
  holidayList.Add CDate("2028/9/22")
  holidayList.Add CDate("2028/10/9")
  holidayList.Add CDate("2028/11/3")
  holidayList.Add CDate("2028/11/23")
  holidayList.Add CDate("2028/12/23")
  holidayList.Add CDate("2029/1/1")
  holidayList.Add CDate("2029/1/8")
  holidayList.Add CDate("2029/2/11")
  holidayList.Add CDate("2029/2/12")
  holidayList.Add CDate("2029/3/20")
  holidayList.Add CDate("2029/4/29")
  holidayList.Add CDate("2029/4/30")
  holidayList.Add CDate("2029/5/3")
  holidayList.Add CDate("2029/5/4")
  holidayList.Add CDate("2029/5/5")
  holidayList.Add CDate("2029/7/16")
  holidayList.Add CDate("2029/8/11")
  holidayList.Add CDate("2029/9/17")
  holidayList.Add CDate("2029/9/23")
  holidayList.Add CDate("2029/9/24")
  holidayList.Add CDate("2029/10/8")
  holidayList.Add CDate("2029/11/3")
  holidayList.Add CDate("2029/11/23")
  holidayList.Add CDate("2029/12/23")
  holidayList.Add CDate("2029/12/24")
  holidayList.Add CDate("2030/1/1")
  holidayList.Add CDate("2030/1/14")
  holidayList.Add CDate("2030/2/11")
  holidayList.Add CDate("2030/3/20")
  holidayList.Add CDate("2030/4/29")
  holidayList.Add CDate("2030/5/3")
  holidayList.Add CDate("2030/5/4")
  holidayList.Add CDate("2030/5/5")
  holidayList.Add CDate("2030/5/6")
  holidayList.Add CDate("2030/7/15")
  holidayList.Add CDate("2030/8/11")
  holidayList.Add CDate("2030/8/12")
  holidayList.Add CDate("2030/9/16")
  holidayList.Add CDate("2030/9/23")
  holidayList.Add CDate("2030/10/14")
  holidayList.Add CDate("2030/11/3")
  holidayList.Add CDate("2030/11/4")
  holidayList.Add CDate("2030/11/23")
  holidayList.Add CDate("2030/12/23")
  holidayList.Add CDate("2031/1/1")
  holidayList.Add CDate("2031/1/13")
  holidayList.Add CDate("2031/2/11")
  holidayList.Add CDate("2031/3/21")
  holidayList.Add CDate("2031/4/29")
  holidayList.Add CDate("2031/5/3")
  holidayList.Add CDate("2031/5/4")
  holidayList.Add CDate("2031/5/5")
  holidayList.Add CDate("2031/5/6")
  holidayList.Add CDate("2031/7/21")
  holidayList.Add CDate("2031/8/11")
  holidayList.Add CDate("2031/9/15")
  holidayList.Add CDate("2031/9/23")
  holidayList.Add CDate("2031/10/13")
  holidayList.Add CDate("2031/11/3")
  holidayList.Add CDate("2031/11/23")
  holidayList.Add CDate("2031/11/24")
  holidayList.Add CDate("2031/12/23")
  holidayList.Add CDate("2032/1/1")
  holidayList.Add CDate("2032/1/12")
  holidayList.Add CDate("2032/2/11")
  holidayList.Add CDate("2032/3/20")
  holidayList.Add CDate("2032/4/29")
  holidayList.Add CDate("2032/5/3")
  holidayList.Add CDate("2032/5/4")
  holidayList.Add CDate("2032/5/5")
  holidayList.Add CDate("2032/7/19")
  holidayList.Add CDate("2032/8/11")
  holidayList.Add CDate("2032/9/20")
  holidayList.Add CDate("2032/9/21")
  holidayList.Add CDate("2032/9/22")
  holidayList.Add CDate("2032/10/11")
  holidayList.Add CDate("2032/11/3")
  holidayList.Add CDate("2032/11/23")
  holidayList.Add CDate("2032/12/23")
End Sub

Public Sub init(ByVal t As Date)
  calDate = t
End Sub


' ---
' Method
' ---
' Return true if is a leap year
Public Function isLeapYear(ByVal y As Integer) As Boolean
  If y Mod 4 = 0 And Not y Mod 100 = 0 Or y Mod 400 = 0 Then
    isLeapYear = True
  Else
    isLeapYear = False
  End If
End Function

' Number of Days in month
Public Function numOfDays() As Integer
  numOfDays = day(lastDayInMonth(calDate))
End Function

' Return the first day in set month.
Public Function firstDayInMonth(ByVal dateVal As Date) As Date
  firstDayInMonth = DateSerial(year(dateVal), month(dateVal), 1)
End Function

' Return the last day in set month.
Public Function lastDayInMonth(ByVal dateVal As Date) As Date
  lastDayInMonth = DateSerial(year(dateVal), month(dateVal) + 1, 0)
End Function

' Return a date list in a month.
Public Function monthlyList() As Object
  Dim dateList As Object
  Dim currentDay As Integer
  Dim startPos As Integer
  Dim endPos As Integer
  Dim wday As Integer
  Dim dateItem As RVDateItem
  Dim i As Variant

  Set dateList = CreateObject("System.Collections.ArrayList")
  currentDay = 1

  ' the position of startday and enddate
  startPos = 0
  Select Case getCalWday(firstDayInMonth(calDate))
  Case 1
    startPos = 0
  Case 2
    startPos = 1
  Case 3
    startPos = 2
  Case 4
    startPos = 3
  Case 5
    startPos = 4
  Case 6
    startPos = 5
  Case Else
    startPos = 6
  End Select
  endPos = (numOfDays() - 1) + startPos

  Dim h As Variant
  For i = 0 To kMonthSize - 1
    Set dateItem = New RVDateItem

    If i < startPos Or i > endPos Then
      dateItem.bgColor = fieldBgColor
      dateItem.color = fieldColor
    Else
      dateItem.dateValue = DateSerial(calYear, calMonth, currentDay)
      wday = getCalWday(dateItem.dateValue)

      ' sat
      If wday = 7 Then
        dateItem.bgColor = saturdayBgColor
        dateItem.color   = saturdayColor
      ElseIf wday = 1 Then
        dateItem.bgColor = offdayBgColor
        dateItem.color   = offdayColor
      ElseIf calYear = year(Date) And calMonth = month(Date) And currentDay = day(Date) Then
        dateItem.bgColor = highlightBgColor
        dateItem.color   = highlightColor
      Else
        dateItem.bgColor = ondayBgColor
        dateItem.color   = ondayColor
      End If

      ' check the day is holiday or not
      If holidayList.contains(dateItem.dateValue) Then
        dateItem.bgColor   = offdayBgColor
        dateItem.color     = offdayColor
        dateItem.isHoliday = True
      End If

      dateItem.value = currentDay
      currentDay     = currentDay + 1
    End If

    dateItem.fontSize = fontSize
    dateList.Add dateItem

    Set dateItem = Nothing
  Next

  Set monthlyList = dateList
  Set dateList    = Nothing
End Function

' Return the weekday val
Public Function getCalWday(ByVal val As Date) As Integer
  getCalWday = Weekday(val)
End Function

Public Function calHead() As Object
  Dim headList As Object
  Dim langList As Object
  Dim dateItem As RVDateItem

  Set headList = CreateObject("System.Collections.ArrayList")
  Set langList = CreateObject("System.Collections.ArrayList")

  Select Case lang
  Case "ja"
    langList.Add "日"
    langList.Add "月"
    langList.Add "火"
    langList.Add "水"
    langList.Add "木"
    langList.Add "金"
    langList.Add "土"
  Case Else
    langList.Add "su"
    langList.Add "mo"
    langList.Add "tu"
    langList.Add "we"
    langList.Add "th"
    langList.Add "fr"
    langList.Add "sa"
  End Select

  Dim k As Variant
  For Each k In langList
    Set dateItem      = New RVDateItem
    dateItem.value    = k
    dateItem.bgColor  = fieldBgColor
    dateItem.color    = fieldColor
    dateItem.fontSize = fontSize

    headList.Add dateItem
    Set dateItem = Nothing
  Next

  Set calHead  = headList
  Set langList = Nothing
  Set headList = Nothing
End Function

' render calendar
Public Sub render(ByVal sheet As Worksheet, ByVal firstRow As Long, ByVal firstColumn As Long, Optional ByVal hasHead As Boolean = True)
  Dim fRow As Long
  Dim fCol As Long
  Dim c As range
  Dim mSize As Integer
  Dim i As Variant

  fRow = firstRow
  fCol = firstColumn

  If fRow < 0 Or fCol < 0 Then
    Exit Sub
  End If

  ' clear
  sheet.Range(Cells(fRow, fCol), Cells(Rows.Count, fCol + kWeekSize)).Clear

  If hasHead Then
    ' print month year
    sheet.Cells(fRow, fCol).value = Format(calDate, "mmm yyyy")
    fRow = fRow + 1

    For i = 1 To kWeekSize
      sheet.Cells(fRow, fCol).ClearContents
      sheet.Cells(fRow, fCol).value          = calHead.item(i - 1).value
      sheet.Cells(fRow, fCol).Font.color     = calHead.item(i - 1).color
      sheet.Cells(fRow, fCol).Font.size      = calHead.item(i - 1).fontSize
      sheet.Cells(fRow, fCol).Interior.color = calHead.item(i - 1).bgColor
      fCol = firstColumn + i
    Next
  End If

  fRow  = fRow + 1
  fCol  = firstColumn
  mSize = monthlyList.Count - 1

  For i = 0 To mSize
    Set c = sheet.Cells(fRow, fCol)

    If monthlyList.item(i).value >= 1 Then
      c.value = monthlyList.item(i).value
    End If

    c.Font.color     = monthlyList.item(i).color
    c.Font.size      = monthlyList.item(i).fontSize
    c.Interior.color = monthlyList.item(i).bgColor

    fCol  = fCol + 1
    Set c = Nothing

    If i Mod kWeekSize = 6 Then
      fCol = firstColumn
      fRow = fRow + 1
    End If
  Next
End Sub
