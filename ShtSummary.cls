VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Dim LastClick As Date
Dim FirstSort As ClsArrow
Dim SecondSort As ClsArrow
Dim Arrows As ClsArrows
Dim ArrowDivision As ClsArrow
Dim ArrowStation As ClsArrow
Dim ArrowAvailability As ClsArrow
Dim ArrowEfficiency As ClsArrow
Dim ClickNo As Integer
Dim StrFirstSort As String
Dim StrSecondSort As String

Private Sub BtnConnectDB_Click()
    ModMain.Initialise
    DBase.DBConnect
End Sub

Private Sub BtnReset_Click()
    ModMain.Initialise
End Sub

Private Sub Worksheet_Activate()
    Dim RecSet As Recordset
    Dim SQL As String
    
    InitialiseArrows

    'get min time value
    SQL = "SELECT MIN(tid) FROM TimeTbl"
    Set RecSet = DBase.SQLQuery(SQL)
    
    'write to sheet
    Range("MinVal").CopyFromRecordset RecSet
    
    'get max time value
    SQL = "SELECT MAX(tid) FROM TimeTbl"
    Set RecSet = DBase.SQLQuery(SQL)
    
    'write to sheet
    Range("MaxVal").CopyFromRecordset RecSet
    
    'set sheet to last month
'    Range("MonthNum") = Range("MaxVal")
    
    'update all data tables
    ShtStationData.GetAvailability Range("MonthNum"), True, True, True, True, [orderby]
    
    'update values
    UpdateSheet
    
    Set RecSet = Nothing
    
End Sub


Private Sub Worksheet_Deactivate()
    TerminateArrows
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim SortBy As DataCol
    Dim SortDir As Direction
        
    If Target.Cells.Count > 1 Then Exit Sub
        
    'toggle help Text
    If Target.Address = "$L$4" Then
        Shapes("Help_dashboard_1").Visible = msoTriStateToggle
        Shapes("Help_dashboard_2").Visible = msoTriStateToggle
        Shapes("Help_dashboard_3").Visible = msoTriStateToggle
        Shapes("Help_dashboard_4").Visible = msoTriStateToggle
        Shapes("Help_dashboard_5").Visible = msoTriStateToggle
        Shapes("Help_dashboard_6").Visible = msoTriStateToggle
        Shapes("Help_dashboard_7").Visible = msoTriStateToggle
        Shapes("Help_dashboard_8").Visible = msoTriStateToggle
        Shapes("Help_dashboard_9").Visible = msoTriStateToggle
        Range("A1").Select
    End If
    
    'prev Month
    If Target.Address = "$V$4" Then
        Range("MonthNum") = ShtStationData.PrevMonth(Range("MonthNum"), Range("MinVal"), Range("MaxVal"))
        UpdateSheet
    End If
    
    'Next Month
    If Target.Address = "$Y$4" Then
        Range("MonthNum") = ShtStationData.NextMonth(Range("MonthNum"), Range("MinVal"), Range("MaxVal"))
        UpdateSheet
    End If
    
    'station details pages
    If Not Intersect(Target, Range("StationDetails")) Is Nothing Then
        Dim Station As String
        Dim StationNo As String
        
        
        'station details button has been pressed.  get station
        StationNo = Target.Offset(0, -10).Value
        Station = Target.Offset(0, -8).Value
        
        'Sort Stations alphabetically first
        DataSort ArwStation, True
        
        'go to station detail sheet and pass station no
        Excel.Range("StationNo") = StationNo
        Excel.Sheets("Station Detail").Activate
        Excel.Range("A1").Select
        
    End If
    
    'sort by Station
    If Target.Address = "$D$4" Then
        DataSort (ArwStation)
    End If
    
    'sort by availability
    If Target.Address = "$E$4" Then
        DataSort (ArwAvailability)
    End If
    
    'sort by Efficiency
    If Target.Address = "$G$4" Then
        DataSort (ArwEfficiency)
    End If
    
    'sort by Division
    If Target.Address = "$C$4" Then
        DataSort (ArwDivision)
    End If
End Sub

Public Sub UpdateSheet()
    Dim InfoGraphic As ClsInfoGraphic
    Dim i As Integer
    Dim Efficiency As Single
    Dim Availability As Single
    Dim Target As Single
    Dim StationLookUp As Range
    Dim Station As Range
    Dim DataSheet As Worksheet
    
    Set DataSheet = ActiveWorkbook.Worksheets("Station Data")
    Set StationLookUp = Range("B5:B42")
    Set InfoGraphic = New ClsInfoGraphic
    
    Set Station = Range("B4")
    For i = 1 To 38
            
        'find station lookup
        Set Station = Station.Offset(1, 0)
        
        'get station data
        Efficiency = Station.Offset(0, 5).Value * 100
        Availability = Station.Offset(0, 3).Value * 100
'        Target = Station.Offset(0, 12).Value * 100
        
        'update graphic
        InfoGraphic.UpdateGraph i, Efficiency, Availability, Target
    Next
    
    ShtSummary.Range("A1").Select
    
    Set DataSheet = Nothing
    Set StationLookUp = Nothing
    Set InfoGraphic = Nothing

End Sub


Public Sub DataSort(SortBy As DataCol, Optional SortAsc As Boolean)
    Dim TimeSinceLastClick As Long
    Dim ActiveArrow As ClsArrow
    Dim StrSortBy As String
    Dim StrDir  As String
       
    'select the correct arrow
    Select Case SortBy
        Case 1
            Set ActiveArrow = ArrowDivision
            StrSortBy = " Division"
        Case 2
            Set ActiveArrow = ArrowStation
            StrSortBy = " SNo"
        Case 3
            Set ActiveArrow = ArrowAvailability
            StrSortBy = " Availability"
        Case 4
            Set ActiveArrow = ArrowEfficiency
            StrSortBy = " Efficiency"
    End Select
    
    'toggle direction
    ActiveArrow.Toggle
    
    If ActiveArrow.Dir = DirUp Or SortAsc Then
        StrDir = " ASC"
    Else
        StrDir = " DESC"
    End If
    
    StrSortBy = StrSortBy & StrDir
    
    'if not the first click, check time since last sort
    If LastClick <> 0 Then
        TimeSinceLastClick = DateDiff("s", LastClick, Now)
    Else
        TimeSinceLastClick = 30
    End If
        
'   Debug.Print  "Last Click: " & LastClick
'    Debug.Print "Time: " & TimeSinceLastClick
    
    If TimeSinceLastClick > 2 Then
    
        'first click - start new sort and reset sort level
        ClickNo = 0
        Arrows.HideAll
        ActiveArrow.Show
        Set FirstSort = ActiveArrow
        StrFirstSort = StrSortBy
        Set SecondSort = Nothing
    Else
    
        'pressed again within time limit so create and manage second sort
        Select Case ClickNo
            Case 0
                If ActiveArrow.DataItem = FirstSort.DataItem Then
                    Set FirstSort = ActiveArrow
                    StrFirstSort = StrSortBy
                Else
                    Set SecondSort = ActiveArrow
                    StrSecondSort = ", " & StrSortBy
                    ActiveArrow.Show
                    ClickNo = 1
                End If
                        
            Case 1
                If ActiveArrow.DataItem = FirstSort.DataItem Then
                    Set FirstSort = ActiveArrow
                    StrFirstSort = StrSortBy
                End If
                If ActiveArrow.DataItem = SecondSort.DataItem Then
                    Set SecondSort = ActiveArrow
                    StrSecondSort = ", " & StrSortBy
                End If
        End Select
    End If
    
'    Debug.Print "Sort level: " & ClickNo
'    Debug.Print "First level: " & FirstSort.DataItem
    If Not SecondSort Is Nothing Then Debug.Print "Second level: " & SecondSort.DataItem
    
    [orderby] = StrFirstSort & StrSecondSort
    
    ShtStationData.GetAvailability Range("MonthNum"), True, True, True, True, [orderby]
    UpdateSheet
    LastClick = Now
    Debug.Print
    
End Sub

Public Sub InitialiseArrows()
    Set Arrows = New ClsArrows
    Set FirstSort = New ClsArrow
    Set SecondSort = New ClsArrow
    Set ArrowAvailability = New ClsArrow
    Set ArrowDivision = New ClsArrow
    Set ArrowEfficiency = New ClsArrow
    Set ArrowStation = New ClsArrow
    
    With ArrowAvailability
        .Arrow = Shapes("Arrow_Availability")
        .Hide
        .Dir = DirUp
        .DataItem = ArwAvailability
        Arrows.AddItem ArrowAvailability
    End With
    
    With ArrowDivision
        .Arrow = Shapes("Arrow_Division")
        .Hide
        .Dir = DirUp
        .DataItem = ArwDivision
        Arrows.AddItem ArrowDivision
    End With
    
    With ArrowEfficiency
        .Arrow = Shapes("Arrow_Efficiency")
        .Hide
        .Dir = DirUp
        .DataItem = ArwEfficiency
        Arrows.AddItem ArrowEfficiency
    End With
    
    With ArrowStation
        .Arrow = Shapes("Arrow_Station")
        .Hide
        .Dir = DirUp
        .DataItem = ArwStation
        Arrows.AddItem ArrowStation
    End With
    
End Sub

Public Sub TerminateArrows()
    Set FirstSort = Nothing
    Set SecondSort = Nothing
    Set Arrows = Nothing
    Set ArrowDivision = Nothing
    Set ArrowStation = Nothing
    Set ArrowAvailability = Nothing
    Set ArrowEfficiency = Nothing

End Sub

