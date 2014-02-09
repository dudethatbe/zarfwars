VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private StartDate As Date
Private CurrentDate As Date
Private CurrentDayInt As Integer
Private DayLimit As Integer
Public Sub Initialize(InitialDate As Date, HowManyDays As Integer)
    StartDate = InitialDate
    DayLimit = HowManyDays
    CurrentDayInt = 1
End Sub
Public Property Get IsUp() As Boolean: IsUp = (CurrentDayInt > DayLimit): End Property
Public Property Get Current_Date() As Date: Current_Date = DateAdd("d", CurrentDayInt, StartDate): End Property
Public Property Get DaysPassed() As Integer: DaysPassed = CurrentDayInt: End Property
Public Sub Crawl(Optional nDays As Integer = 1)
    [d1].Columns.AutoFit
    [d1] = Current_Date
    [d2] = DayLimit - CurrentDayInt
    CurrentDayInt = CurrentDayInt + 1
End Sub
