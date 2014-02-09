VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZarfClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private zName As String
Private zEffect As String
Private zPrices() As Variant
Private HomeRange As Range
Public Property Get Name() As String: Name = zName: End Property
Public Property Let Name(newName As String): zName = newName: End Property
Public Property Get Effect() As String: Effect = zEffect: End Property
Public Property Let Effect(newEffect As String): zEffect = newEffect: End Property
Public Property Get Cell_Location() As Range: Set Cell_Location = HomeRange: End Property
Public Property Get Prices() As Variant(): Prices = zPrices: End Property
Public Sub Initialize(dest As Range, zarf_name As String, zarf_effect As String, ParamArray Zarf_Prices() As Variant)
    Set HomeRange = dest
    Me.Name = zarf_name & " Zarf"
    zEffect = zarf_effect
    zPrices = Zarf_Prices
End Sub
Private Sub Class_Terminate()
    Erase zPrices
    zName = ""
End Sub
Public Function Pick_Price() As String
    Pick_Price = "$" & zPrices(Rnd * UBound(zPrices))
End Function
