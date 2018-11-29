VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrewData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

End Sub



Public Sub GetCrewData()
    Dim RstCrewData As Recordset
    Dim DataRange As Range
    Dim StationID As String
    Dim MonthNo As Integer
    
    StationID = "'" & Globals.ConvStnNotoID(StationDetail.Range("StationNo")) & "'"
    MonthNo = Summary.Range("Monthnum")
    
    ClearData
    
    Set DataRange = Range("A2")
    
    Set RstCrewData = DBase.SQLQuery("SELECT * FROM Crewmember WHERE Tid = " & MonthNo & " AND Station = " & StationID & "ORDER BY totalhrs DESC")
    
    'write values
    DataRange.CopyFromRecordset RstCrewData
    
End Sub

Public Sub ClearData()
    Range("A2:K50").ClearContents
    
End Sub
