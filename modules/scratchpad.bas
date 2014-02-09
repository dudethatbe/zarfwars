Attribute VB_Name = "scratchpad"
Public Sub randomnumber()
    Debug.Print Int(Rnd() * 4)
End Sub
Public Sub zarfem()
Dim WhiteZarf As ZarfClass
    Set WhiteZarf = New ZarfClass
    WhiteZarf.Name = "White"
    WhiteZarf.Set_Prices 12, 18, 21
    Debug.Print WhiteZarf.Pick_Price
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
Private Sub VariablePrices()

End Sub
