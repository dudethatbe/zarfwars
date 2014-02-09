Attribute VB_Name = "Main"
Option Explicit
Dim Location As LocationClass
Dim Time As TimeClass
Dim ZarfColl As Collection
Dim FirstDay As Date
Dim EarlyQuit As Boolean
Public Sub Run()
If Not Main_Init Then
    Exit Sub
Else
    Do
        Handle_Events
        Time.Crawl
    Loop Until Time.IsUp Or EarlyQuit
End If
Zarf_Cleanup
End Sub

Public Sub Handle_Events()
Dim UserResponse As VbMsgBoxResult
    ' update date
    [d1] = DateAdd("d", Time.DaysPassed, FirstDay)
    [d1].Columns.AutoFit
    [d2] = [d2] - 1
    Location.RandPick
    UserResponse = MsgBox("Currently in " & Location.CurrentLocation & vbCrLf & "Continue Zarfing?", vbYesNo, "ZarfWars")
    If UserResponse = vbNo Then
        EarlyQuit = True
    End If
End Sub
Public Function Main_Init() As Boolean
Dim j As Range
    Set Location = New LocationClass
    If Location.CurrentLocation = "" Then
        Main_Init = False
    Else
        Set Time = New TimeClass
        Time.Initialize "12/8/1989", 30
        MakeSheet "ZarfWars"
        FirstDay = "12/8/1989"
        [a1] = "Debt:"
        [a2] = "Cash:"
        [b1] = "$5,000"
        [b2] = "$200"
        [c1] = "Date:"
        [c2] = "Days Left:"
        [d1] = FirstDay
        [d2] = "30"
        [a4] = "Locations"
        BoldRanges "a1", "a2", "c1", "c2", "a4", "b5", "c5", "d5", "e5"
        Range("b4", "e4").Value = Location.AllLocations
        SetZarfs
        EarlyQuit = False
        Main_Init = True
    End If
End Function
Private Sub Zarf_Cleanup()
    Set Time = Nothing
    Set Location = Nothing
End Sub
Private Sub SetZarfs()
Dim ZarfColl As Collection
Dim zNames() As Variant: zNames = Array("Green", "Yellow", "Blue", "Brown")
Dim zEffects() As Variant: zEffects = Array("!!!", "zzz", "???", "~~~")
Dim zPrices() As Variant: zPrices = Array(Array(100, 85, 50), Array(80, 70, 45), Array(40, 20, 15), Array(15, 8, 3))
Dim Zarf As ZarfClass
Dim i As Integer

    Set ZarfColl = New Collection
    For i = 0 To 3
        Set Zarf = New ZarfClass
        Zarf.Initialize Cells(i + 2, 6), CStr(zNames(i)), CStr(zEffects(i)), zPrices(i)(0), zPrices(i)(1), zPrices(i)(2)
        ZarfColl.Add Zarf, Zarf.Name
    Next i
    Set Zarf = Nothing
    For i = 1 To ZarfColl.Count
        With ZarfColl.Item(i)
            .Cell_Location = .Name
            .Cell_Location.Offset(0, 1) = .Effect
            .Cell_Location.Columns.AutoFit
        End With
    Next i
    Set Zarf = Nothing
    Set ZarfColl = Nothing
End Sub
Public Sub BoldRanges(ParamArray ranges() As Variant)
Dim j As Variant
    For Each j In ranges
        Range(j).Font.Bold = True
    Next j
End Sub
Private Sub MakeSheet(x As String)
Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(x)
    If Not ws Is Nothing Then
        With Range("a1", "e5")
            .Font.Bold = False
            .Value = ""
        End With
    Else
        Sheets.Add
        Set ws = ActiveSheet
        ws.Name = x
    End If
    On Error GoTo 0
End Sub
