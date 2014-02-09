VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private AllLoc() As String
Private CurLoc As String
Public Enum Location_Set
    The_Mall
    Parkland
    Sixth_Ave
    Hilltop
End Enum
Public Property Get CurrentLocation() As String: CurrentLocation = CurLoc: End Property
Public Property Get AllLocations() As String(): AllLocations = AllLoc: End Property
Private Sub Class_Initialize()
    AllLoc = Split("The Mall,Parkland,6th Ave,Hilltop", ",")
    CurLoc = Location_Str(Rnd() * 4)
End Sub
Public Sub RandPick()
Dim temp As Range
    CurLoc = Location_Str(Rnd() * 4)
    Range("b5", "e5").Value = ""
    Select Case CurLoc
        Case Is = "The Mall"
            Set temp = [b5]
        Case Is = "Parkland"
            Set temp = [c5]
        Case Is = "6th Ave"
            Set temp = [d5]
        Case Is = "Hilltop"
            Set temp = [e5]
    End Select
    temp.Value = "^"
Set temp = Nothing
End Sub
Private Sub Class_Terminate()
    Erase AllLoc
End Sub
Private Function Location_Str(n As Integer) As String
' function converts location_set type to string
    If n < 4 And n > -1 Then
        Location_Str = AllLoc(n)
    Else 'default hilltop
        Location_Str = "Hilltop"
    End If
End Function
