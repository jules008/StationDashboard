VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pStationNo As String
Private pTid As Integer
Private pStatus As String
Private pACA As Boolean
Private pEfficiency As Single
Private pAvailability As Single
Private pAvail_ACA As Boolean
Private pAvail_Ext As Single
Private pAvailTarget As Single
Private pNumActiveCrew As Integer
Private pTotalHrs As Integer
Private pOTR_LowCrew As Single
Private pOTR_NoCriver As Single
Private pOTR_NoJO As Single
Private pCrewMembers As Collection

' StationNo property
'------------------
Public Property Get StationNo() As String
    StationNo = pStationNo
End Property
Public Property Let StationNo(Value As String)
    pStationNo = Value
End Property

' Tid property
'------------------
Public Property Get Tid() As Integer
    Tid = pTid
End Property
Public Property Let Tid(Value As Integer)
    pTid = Value
End Property

' Status property
'------------------
Public Property Get Status() As String
    Status = pStatus
End Property
Public Property Let Status(Value As String)
    pStatus = Value
End Property

' Efficiency property
'------------------
Public Property Get Efficiency() As Single
    Efficiency = pEfficiency
End Property
Public Property Let Efficiency(Value As Single)
    pEfficiency = Value
End Property

' ACA property
'------------------
Public Property Get ACA() As Boolean
    ACA = pACA
End Property
Public Property Let ACA(Value As Boolean)
    pACA = Value
End Property

' Availability property
'------------------
Public Property Get Availability() As Single
    Availability = pAvailability
End Property
Public Property Let Availability(Value As Single)
    pAvailability = Value
End Property

' Avail_ACA property
'------------------
Public Property Get Avail_ACA() As Single
    Avail_ACA = pAvail_ACA
End Property
Public Property Let Avail_ACA(Value As Single)
    pAvail_ACA = Value
End Property

' Avail_Ext property
'------------------
Public Property Get Avail_Ext() As Single
    Avail_Ext = pAvail_Ext
End Property
Public Property Let Avail_Ext(Value As Single)
    pAvail_Ext = Value
End Property

' AvailTarget property
'------------------
Public Property Get AvailTarget() As Single
    AvailTarget = pAvailTarget
End Property
Public Property Let AvailTarget(Value As Single)
    pAvailTarget = Value
End Property

' NumActiveCrew property
'------------------
Public Property Get NumActiveCrew() As Integer
    NumActiveCrew = pNumActiveCrew
End Property
Public Property Let NumActiveCrew(Value As Integer)
    pNumActiveCrew = Value
End Property

' TotalHrs property
'------------------
Public Property Get TotalHrs() As Integer
    TotalHrs = pTotalHrs
End Property
Public Property Let TotalHrs(Value As Integer)
    pTotalHrs = Value
End Property

' OTR_LowCrew property
'------------------
Public Property Get OTR_LowCrew() As Single
    OTR_LowCrew = pOTR_LowCrew
End Property
Public Property Let OTR_LowCrew(Value As Single)
    pOTR_LowCrew = Value
End Property

' OTR_NoCriver property
'------------------
Public Property Get OTR_NoCriver() As Single
    OTR_NoCriver = pOTR_NoCriver
End Property
Public Property Let OTR_NoCriver(Value As Single)
    pOTR_NoCriver = Value
End Property

' OTR_NoJO property
'------------------
Public Property Get OTR_NoJO() As Single
    OTR_NoJO = pOTR_NoJO
End Property
Public Property Let OTR_NoJO(Value As Single)
    pOTR_NoJO = Value
End Property

'======================================================================
'Methods
'======================================================================

'Count Method
'------------------
Public Property Get Count() As Long
    Count = pCrewMembers.Count
End Property

'Count drivers Method
'------------------
Public Property Get DriverCount() As Integer
    Dim Count As Integer
    
    For Each CCrewMember In pCrewMembers
        If CCrewMember.DRIVER = True Then Count = Count + 1
    Next
    DriverCount = Count
End Property

'Count OIC Method
'------------------
Public Property Get OICCount() As Integer
    Dim Count As Integer
    
    For Each CCrewMember In pCrewMembers
        If CCrewMember.OIC = True Then Count = Count + 1
    Next
    OICCount = Count
End Property

'Count STC Method
'------------------
Public Property Get STCCount() As Integer
    Dim Count As Integer
    
    For Each CCrewMember In pCrewMembers
        If CCrewMember.STC = True Then Count = Count + 1
    Next
    STCCount = Count
End Property



'Remove Method
'------------------
Public Function Remove(IndexorName As String)
    pCrewMembers.Remove (IndexorName)
End Function

'List Method
'------------------
Public Function List()
    Dim i As Integer
    
    For Each CCrewMember In CStations
        Debug.Print CCrewMember.CrewNo
        Debug.Print CCrewMember.EventTime
        Debug.Print CCrewMember.OnOffCall
        Debug.Print CCrewMember.Station
        Debug.Print
        
    Next
End Function


Public Function Find(CrewNo As String) As Integer
    Dim i As Integer
    
    For Each CCrewMember In pCrewMembers
        i = i + 1
        
        If CCrewMember.CrewNo = CrewNo Then
            Find = i
            Exit Function
        End If
    Next
    
    Find = 0
    
End Function

Private Sub Class_Initialize()
    Set pCrewMembers = New Collection
End Sub
