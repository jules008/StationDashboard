VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtStationDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim PencilGraph1 As ClsPencilGraphic
Dim PencilGraph2 As ClsPencilGraphic
Dim PencilGraph3 As ClsPencilGraphic
Dim PencilGraph4 As ClsPencilGraphic
Dim PencilGraph5 As ClsPencilGraphic
Dim PencilGraph6 As ClsPencilGraphic
Dim PencilGraph7 As ClsPencilGraphic
Dim PencilGraphs As ClsPencilGraphics

Enum PencilGraphType
    NoCrew = 1
    NoDrivers = 2
    NoJOs = 3
    HoursBooked = 4
End Enum

Private Sub CommandButton1_Click()
    ModMain.Initialise

End Sub

Private Sub OptACA_Click()
    Range("ACAExtCovCol").Value = "ACA"
End Sub

Private Sub OptExtCov_Click()
    Range("ACAExtCovCol").Value = "Ext"
End Sub

Private Sub OptionButton1_Click()
    Dim RngGraphType As Range
    Dim GraphType As PencilGraphType
    Dim KeyCrew As Shape
    Dim KeyDrJO As Shape
    
    Set KeyCrew = Shapes("key_crew")
    Set KeyDrJO = Shapes("Key_drJO")
    
    KeyCrew.Visible = msoCTrue
    KeyDrJO.Visible = msoFalse
    
    Me.Unprotect StrPass
    Set RngGraphType = Range("graphType")
    GraphType = NoCrew
    RngGraphType = GraphType
    SetThresholds [GraphType]
    UpdatePencilGraphic
    Me.Protect StrPass
    
    Set KeyCrew = Nothing
    Set KeyDrJO = Nothing

End Sub

Private Sub OptionButton2_Click()
    Dim RngGraphType As Range
    Dim GraphType As PencilGraphType
    Dim KeyCrew As Shape
    Dim KeyDrJO As Shape
    
    Set KeyCrew = Shapes("key_crew")
    Set KeyDrJO = Shapes("Key_drJO")
    
    KeyCrew.Visible = msoFalse
    KeyDrJO.Visible = msoCTrue
    
    Me.Unprotect StrPass
    Set RngGraphType = Range("graphType")
    GraphType = NoDrivers
    RngGraphType = GraphType
    SetThresholds [GraphType]
    UpdatePencilGraphic
    
    Range("A1").Select
    Me.Protect StrPass
    
    Set KeyCrew = Nothing
    Set KeyDrJO = Nothing
    
End Sub

Private Sub OptionButton3_Click()
    Dim RngGraphType As Range
    Dim GraphType As PencilGraphType
    Dim KeyCrew As Shape
    Dim KeyDrJO As Shape
    
    Set KeyCrew = Shapes("key_crew")
    Set KeyDrJO = Shapes("Key_drJO")
    
    KeyCrew.Visible = msoFalse
    KeyDrJO.Visible = msoCTrue
    
    
    Set RngGraphType = Range("graphType")
    Me.Unprotect StrPass
    GraphType = NoJOs
    RngGraphType = GraphType
    SetThresholds [GraphType]
    UpdatePencilGraphic
    Me.Protect StrPass
    
    Set KeyCrew = Nothing
    Set KeyDrJO = Nothing

End Sub

Private Sub OptTotal_Click()
    Range("ACAExtCovCol").Value = "Total"
End Sub

Private Sub SpnWeek_Change()
    [weekno] = SpnWeek.Value
    UpdatePencilGraphic
End Sub

Private Sub Worksheet_Activate()
    Me.OptTotal.Value = True
    Range("ACAExtCovCol").Value = "Total"
    
    UpdateAvailGraphic
    
    ShtCrewData.GetCrewData
    
    Set PencilGraphs = New ClsPencilGraphics
    Set PencilGraph1 = New ClsPencilGraphic
    Set PencilGraph2 = New ClsPencilGraphic
    Set PencilGraph3 = New ClsPencilGraphic
    Set PencilGraph4 = New ClsPencilGraphic
    Set PencilGraph5 = New ClsPencilGraphic
    Set PencilGraph6 = New ClsPencilGraphic
    Set PencilGraph7 = New ClsPencilGraphic
    
    With PencilGraphs
        .AddItem PencilGraph1
        .AddItem PencilGraph2
        .AddItem PencilGraph3
        .AddItem PencilGraph4
        .AddItem PencilGraph5
        .AddItem PencilGraph6
        .AddItem PencilGraph7
    End With
    
    For i = 1 To 48
        PencilGraph1.Shapes.Add Me.Shapes("Mon" & i)
        PencilGraph2.Shapes.Add Me.Shapes("Tue" & i)
        PencilGraph3.Shapes.Add Me.Shapes("Wed" & i)
        PencilGraph4.Shapes.Add Me.Shapes("Thu" & i)
        PencilGraph5.Shapes.Add Me.Shapes("Fri" & i)
        PencilGraph6.Shapes.Add Me.Shapes("Sat" & i)
        PencilGraph7.Shapes.Add Me.Shapes("Sun" & i)
    Next
    
    SetThresholds [GraphType]
    UpdatePencilGraphic
    OptionButton1.Value = True
    
End Sub

Private Sub Worksheet_Deactivate()
    Set PencilGraph1 = Nothing
    Set PencilGraph2 = Nothing
    Set PencilGraph3 = Nothing
    Set PencilGraph4 = Nothing
    Set PencilGraph5 = Nothing
    Set PencilGraph6 = Nothing
    Set PencilGraph7 = Nothing
    Set PencilGraphs = Nothing
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    If Target.Cells.Count > 1 Then Exit Sub
        
    'month scroll
    If Target.Address = "$P$1" Then
        
        'next month
        With Summary
            .Range("MonthNum").Value = StationData.NextMonth(.Range("MonthNum"), .Range("MinVal"), .Range("MaxVal"))
        End With
        UpdateAvailGraphic
        UpdatePencilGraphic
        ShtCrewData.GetCrewData
        Range("A1").Select
 
    End If
    
    If Target.Address = "$M$1" Then
        With Summary
            .Range("MonthNum").Value = StationData.PrevMonth(.Range("MonthNum"), .Range("MinVal"), .Range("MaxVal"))
        End With
        UpdateAvailGraphic
        UpdatePencilGraphic
        ShtCrewData.GetCrewData
        Range("A1").Select
    End If
    
    'station scroll
    If Target.Address = "$B$1" Then
        
        StationNo = Range("StationNo")
        
        If StationNo > 1 Then
            StationNo = StationNo - 1
            Range("StationNo") = StationNo
            UpdateAvailGraphic
            UpdatePencilGraphic
            ShtCrewData.GetCrewData
        End If
        
        Range("A1").Select
        
    End If
    
    If Target.Address = "$F$1" Then
        
        StationNo = Range("StationNo")
        
        If StationNo < 38 Then
            StationNo = StationNo + 1
            Range("StationNo") = StationNo
            UpdateAvailGraphic
            UpdatePencilGraphic
            ShtCrewData.GetCrewData
        End If
        Range("A1").Select
        
    End If
    
        
    'help text
    If Target.Address = "$B$3" Then
        Shapes("Help_Avail_1").Visible = msoTriStateToggle
        Shapes("Help_Avail_2").Visible = msoTriStateToggle
        Shapes("Help_Avail_3").Visible = msoTriStateToggle
        Shapes("Help_Avail_4").Visible = msoTriStateToggle
        Range("A1").Select
    End If
        
    If Target.Address = "$B$20" Then
        Shapes("Help_Cover_1").Visible = msoTriStateToggle
        Shapes("Help_Cover_2").Visible = msoTriStateToggle
        Range("A1").Select
    End If
    
    If Target.Address = "$M$3" Then
        Shapes("Help_Crew_1").Visible = msoTriStateToggle
        Shapes("Help_Crew_2").Visible = msoTriStateToggle
        Range("A1").Select
    End If
    
    If Target.Address = "$U$3" Then
        Shapes("Help_Trend_1").Visible = msoTriStateToggle
        Shapes("Help_Trend_2").Visible = msoTriStateToggle
        Shapes("Help_Trend_3").Visible = msoTriStateToggle
        Shapes("Help_Trend_4").Visible = msoTriStateToggle
        Shapes("Help_Trend_5").Visible = msoTriStateToggle
        Range("A1").Select
    End If
End Sub

Private Sub UpdateAvailGraphic()
    Dim Background As Shape
    Dim ShpNetCover As Shape
    Dim ShpACACover As Shape
    Dim ShpExtCover As Shape
    Dim ShpTgt As Shape
    Dim AvailTgt As Single
    Dim AvailNet As Single
    Dim AvailACA As Single
    Dim AvailExt As Single
    
    Set Background = Shapes("ShpAvailBack")
    Set ShpNetCover = Shapes("ShpNetAvail")
    Set ShpACACover = Shapes("ShpAvailACA")
    Set ShpExtCover = Shapes("ShpAvailExt")
    Set ShpTgt = Shapes("ShpAvailTgt")
    
    AvailTgt = Range("G15")
    AvailNet = Range("G19")
    AvailACA = AvailNet + Range("G17")
    AvailExt = AvailACA + Range("G18")
    
    Background.Width = 250
    ShpExtCover.Width = AvailExt * 250
    ShpACACover.Width = AvailACA * 250
    ShpNetCover.Width = AvailNet * 250
    ShpTgt.Left = 35 + (AvailTgt * 250)
  
    
    
    
    Set Background = Nothing
    Set ShpNetCover = Nothing
    Set ShpACACover = Nothing
    Set ShpExtCover = Nothing
    
    
End Sub

Private Sub SetThresholds(GraphType As PencilGraphType)
    Dim i As Integer
    Dim Threshold1 As Integer
    Dim Threshold2 As Integer
    Dim Threshold3 As Integer
    Dim Threshold4 As Integer
    Dim Threshold5 As Integer
    Dim Threshold6 As Integer
    Dim Colour1 As Long
    Dim Colour2 As Long
    Dim Colour3 As Long
    Dim Colour4 As Long
    Dim Colour5 As Long
    Dim Colour6 As Long
    Dim Colour7 As Long
    Dim LocalGraph As ClsPencilGraphic
    
    Set LocalGraph = New ClsPencilGraphic
    
    Select Case GraphType
        Case NoCrew
            Threshold1 = 2
            Threshold2 = 3
            Threshold3 = 4
            Threshold4 = 5
            Threshold5 = 6
            Threshold6 = 7
            Threshold7 = 8
            Threshold8 = 9
            Threshold9 = 10
            Colour1 = RGB(28, 17, 10)
            Colour2 = RGB(255, 242, 147)
            Colour3 = RGB(255, 200, 87)
            Colour4 = RGB(160, 230, 135)
            Colour5 = RGB(130, 209, 115)
            Colour6 = RGB(100, 180, 94)
            Colour7 = RGB(68, 150, 74)
            Colour8 = RGB(58, 125, 63)
            Colour9 = RGB(48, 107, 52)
        Case NoDrivers
            Threshold1 = 1
            Threshold2 = 2
            Threshold3 = 3
            Threshold4 = 4
            Threshold5 = 5
            Threshold6 = 6
            Threshold7 = 7
            Threshold8 = 8
            Threshold9 = 9
            Colour1 = RGB(28, 17, 10)
            Colour2 = RGB(160, 230, 135)
            Colour3 = RGB(130, 209, 115)
            Colour4 = RGB(100, 180, 94)
            Colour5 = RGB(68, 150, 74)
            Colour6 = RGB(68, 150, 74)
            Colour7 = RGB(68, 150, 74)
            Colour8 = RGB(68, 150, 74)
            Colour9 = RGB(68, 150, 74)
        Case NoJOs
            Threshold1 = 1
            Threshold2 = 2
            Threshold3 = 3
            Threshold4 = 4
            Threshold5 = 5
            Threshold6 = 6
            Threshold7 = 7
            Threshold8 = 8
            Threshold9 = 9
            Colour1 = RGB(28, 17, 10)
            Colour2 = RGB(160, 230, 135)
            Colour3 = RGB(130, 209, 115)
            Colour4 = RGB(100, 180, 94)
            Colour5 = RGB(68, 150, 74)
            Colour6 = RGB(68, 150, 74)
            Colour7 = RGB(68, 150, 74)
            Colour8 = RGB(68, 150, 74)
            Colour9 = RGB(68, 150, 74)
        Case HoursBooked
            Threshold1 = 1
            Threshold2 = 2
            Threshold3 = 3
            Threshold4 = 4
            Threshold5 = 5
            Threshold6 = 6
            Threshold7 = 7
            Colour1 = RGB(28, 17, 10)
            Colour2 = RGB(160, 230, 135)
            Colour3 = RGB(130, 209, 115)
            Colour4 = RGB(100, 180, 94)
            Colour5 = RGB(68, 150, 74)
            Colour6 = RGB(58, 125, 63)
            Colour7 = RGB(48, 107, 52)
            Colour8 = RGB(58, 125, 63)
            Colour9 = RGB(48, 107, 52)
    End Select
    
    For i = 1 To 7
        Set LocalGraph = PencilGraphs.FindItem(i)
                    
        With LocalGraph
            .Threshold1 = Threshold1
            .Threshold2 = Threshold2
            .Threshold3 = Threshold3
            .Threshold4 = Threshold4
            .Threshold5 = Threshold5
            .Threshold6 = Threshold6
            .Threshold7 = Threshold7
            .Threshold8 = Threshold8
            .Threshold9 = Threshold9
            .Threshold1Col = Colour1
            .Threshold2Col = Colour2
            .Threshold3Col = Colour3
            .Threshold4Col = Colour4
            .Threshold5Col = Colour5
            .Threshold6Col = Colour6
            .Threshold7Col = Colour7
            .Threshold8Col = Colour8
            .Threshold9Col = Colour9
        End With
    Next
    Set LocalGraph = Nothing
End Sub


Public Sub UpdatePencilGraphic()
    Dim VarInputValues() As Variant
    Dim DataType As PencilGraphType
    Dim DateNo As Integer
    Dim MonthNo As Integer
    Dim StationNo As Integer
    
    MonthNo = [MonthNum]
    StationNo = Range("StationNo")
    DataType = [GraphType]
    DateNo = GetStartDay
    
    Debug.Print "DateNo: " & DateNo
    
    PencilGraph1.Update StationNo, MonthNo, DateNo, DataType
    PencilGraph2.Update StationNo, MonthNo, DateNo + 1, DataType
    PencilGraph3.Update StationNo, MonthNo, DateNo + 2, DataType
    PencilGraph4.Update StationNo, MonthNo, DateNo + 3, DataType
    PencilGraph5.Update StationNo, MonthNo, DateNo + 4, DataType
    PencilGraph6.Update StationNo, MonthNo, DateNo + 5, DataType
    PencilGraph7.Update StationNo, MonthNo, DateNo + 6, DataType
    

End Sub

Public Function GetStartDay() As Integer
    Dim StartDay As Integer
    
    StartDay = Summary.Range("StartDay")
    
    If StartDay > 1 Then
        StartDay = 2 - StartDay
    End If
    
    StartDay = StartDay + ([weekno] - 1) * 7
    
    GetStartDay = StartDay
    
End Function

