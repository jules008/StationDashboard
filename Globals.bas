Attribute VB_Name = "Globals"
Option Explicit

Public DBase As ClsDatabase
Public Const StrPass = "moowej"
Enum Direction
    DirUp = 1
    DirDown = 2
End Enum

Enum DataCol
    ArwDivision = 1
    ArwStation = 2
    ArwAvailability = 3
    ArwEfficiency = 4
End Enum

Public Function ConvStnNotoID(StationNo As Integer) As String
    ConvStnNotoID = "EC" & CStr(Format(StationNo, "00"))

End Function
