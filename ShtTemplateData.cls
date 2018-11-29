VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtTemplateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Public Sub RefreshData()
    Dim SQL As String
    Dim RstTemplateData As Recordset
    Dim RstCrewNo As Recordset
    Dim TargetRange As Range
    
    ClearSheet
    
    Set TargetRange = ShtTemplateData.Range("A2")
    
    Set RstTemplateData = DBase.SQLQuery("SELECT TemplateDetail.CrewNo, TemplateDetail.StationNo, TemplateDetail.OnCall, Template.NoStation, Template.StationNo, Count(TemplateDetail.CrewNo) AS CountOfCrewNo" & _
                                            " FROM Template RIGHT JOIN TemplateDetail ON (Template.NoStation = TemplateDetail.StationNo) AND (Template.CrewNo = TemplateDetail.CrewNo)" & _
                                            " GROUP BY TemplateDetail.CrewNo, TemplateDetail.StationNo, TemplateDetail.OnCall, Template.NoStation, Template.StationNo" & _
                                            " ORDER BY TemplateDetail.CrewNo")

    TargetRange.CopyFromRecordset RstTemplateData
    
    Set RstCrewNo = DBase.SQLQuery("SELECT DISTINCT crewno FROM template")
    
    TargetRange.Offset(0, 7).CopyFromRecordset RstCrewNo
    
    
    
End Sub

Public Sub ClearSheet()
    Range("A2:H2000").ClearContents
End Sub
