VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPencilGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,01 - Initial version
'---------------------------------------------------------------
' Date - 26 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' LayoutGraphic - Sets out Pencil Graphic
' DeleteShapes - deletes all shapes
' UpdateGraphic - Updates the graphic
' ReturnBanding - Returns the numerical value of the banding
'===============================================================
Option Explicit

Private pPosLeft As Integer
Private pPosTop As Integer
Private pTotHeight As Integer
Private pTotWidth As Integer
Private pThreshold1 As Variant
Private pThreshold1Col As Long
Private pThreshold2 As Variant
Private pThreshold2Col  As Long
Private pThreshold3 As Variant
Private pThreshold3Col As Long
Private pThreshold4 As Variant
Private pThreshold4Col As Long
Private pThreshold5 As Variant
Private pThreshold5Col As Long
Private pThreshold6 As Variant
Private pThreshold6Col As Long
Private pThreshold7 As Variant
Private pThreshold7Col As Long
Private pThreshold8 As Variant
Private pThreshold8Col As Long
Private pThreshold9 As Variant
Private pThreshold9Col As Long
Private pShapes As Collection


'---------------------------------------------------------------
Public Property Get PosLeft() As Integer
    PosLeft = pPosLeft
End Property

Public Property Let PosLeft(ByVal vNewValue As Integer)
    pPosLeft = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PosTop() As Integer
    PosTop = pPosTop
End Property

Public Property Let PosTop(ByVal vNewValue As Integer)
   pPosTop = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TotHeight() As Integer
    TotHeight = pTotHeight
End Property

Public Property Let TotHeight(ByVal vNewValue As Integer)
     pTotHeight = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TotWidth() As Integer
    TotWidth = pTotWidth
End Property

Public Property Let TotWidth(ByVal vNewValue As Integer)
     pTotWidth = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold1() As Variant
    Threshold1 = pThreshold1
End Property

Public Property Let Threshold1(ByVal vNewValue As Variant)
     pThreshold1 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold1Col() As Long
    Threshold1Col = pThreshold1Col
End Property

Public Property Let Threshold1Col(ByVal vNewValue As Long)
    pThreshold1Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold2() As Variant
    Threshold2 = pThreshold2
End Property

Public Property Let Threshold2(ByVal vNewValue As Variant)
     pThreshold2 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold2Col() As Long
    Threshold2Col = pThreshold2Col
End Property

Public Property Let Threshold2Col(ByVal vNewValue As Long)
     pThreshold2Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold3() As Variant
    Threshold3 = pThreshold3
End Property

Public Property Let Threshold3(ByVal vNewValue As Variant)
     pThreshold3 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold3Col() As Long
    Threshold3Col = pThreshold3Col
End Property

Public Property Let Threshold3Col(ByVal vNewValue As Long)
     pThreshold3Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold4() As Variant
    Threshold4 = pThreshold4
End Property

Public Property Let Threshold4(ByVal vNewValue As Variant)
    pThreshold4 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold4Col() As Long
    Threshold4Col = pThreshold4Col
End Property

Public Property Let Threshold4Col(ByVal vNewValue As Long)
     pThreshold4Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold5() As Variant
    Threshold5 = pThreshold5
End Property

Public Property Let Threshold5(ByVal vNewValue As Variant)
     pThreshold5 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold5Col() As Long
    Threshold5Col = pThreshold5Col
End Property

Public Property Let Threshold5Col(ByVal vNewValue As Long)
     pThreshold5Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold6() As Variant
    Threshold6 = pThreshold6
End Property

Public Property Let Threshold6(ByVal vNewValue As Variant)
     pThreshold6 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold6Col() As Long
    Threshold6Col = pThreshold6Col
End Property

Public Property Let Threshold6Col(ByVal vNewValue As Long)
     pThreshold6Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold7() As Variant
    Threshold7 = pThreshold7
End Property

Public Property Let Threshold7(ByVal vNewValue As Variant)
     pThreshold7 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold7Col() As Long
    Threshold7Col = pThreshold7Col
End Property

Public Property Let Threshold7Col(ByVal vNewValue As Long)
     pThreshold7Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold8() As Variant
    Threshold8 = pThreshold8
End Property

Public Property Let Threshold8(ByVal vNewValue As Variant)
     pThreshold8 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold8Col() As Long
    Threshold8Col = pThreshold8Col
End Property

Public Property Let Threshold8Col(ByVal vNewValue As Long)
     pThreshold8Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold9() As Variant
    Threshold9 = pThreshold9
End Property

Public Property Let Threshold9(ByVal vNewValue As Variant)
     pThreshold9 = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Threshold9Col() As Long
    Threshold9Col = pThreshold9Col
End Property

Public Property Let Threshold9Col(ByVal vNewValue As Long)
     pThreshold9Col = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Shapes() As Collection
    Set Shapes = pShapes
End Property

Public Property Let Shapes(ByVal vNewValue As Collection)
     Set pShapes = Shapes
End Property

'===============================================================
' Method DeleteShapes - Deletes all shapes
'===============================================================
Public Sub DeleteShapes()
    Dim Shp As Shape
    
    For Each Shp In pShapes
        Shp.Delete
    Next
End Sub

'===============================================================
' Method ReturnBanding - Returns the numerical value of the banding
'===============================================================
Private Function ReturnBanding(InputValue As Variant) As Long
    Select Case InputValue
        Case Is < pThreshold1
            ReturnBanding = pThreshold1Col
        Case Is < pThreshold2
            ReturnBanding = pThreshold2Col
        Case Is < pThreshold3
            ReturnBanding = pThreshold3Col
        Case Is < pThreshold4
            ReturnBanding = pThreshold4Col
        Case Is < pThreshold5
            ReturnBanding = pThreshold5Col
        Case Is < pThreshold6
            ReturnBanding = pThreshold6Col
        Case Is < pThreshold7
            ReturnBanding = pThreshold7Col
        Case Is < pThreshold8
            ReturnBanding = pThreshold8Col
        Case Else
            ReturnBanding = pThreshold9Col
    End Select
    
End Function

'===============================================================
' Method Update - Updates the Pencil Graphic from StationNo,
'                 MonthNo and DateNo
'===============================================================
Public Sub Update(StationNo As Integer, MonthNo As Integer, DateNo As Integer, DataType As PencilGraphType)
    Dim Shp As Shape
    Dim i As Integer
    Dim InputData() As Variant
    
    InputData = GetData(MonthNo, StationNo, DateNo, DataType)
    
    i = 1
    For Each Shp In pShapes
        Shp.Fill.ForeColor.RGB = ReturnBanding(InputData(i))
        i = i + 1
    Next

End Sub

'===============================================================
' Method GetData - Gets data from database
'===============================================================
Private Function GetData(MonthNo As Integer, StationNo As Integer, DateNo As Integer, DataType As PencilGraphType) As Variant
    Dim RstPencilData As Recordset
    Dim StationID As String
    Dim PrevData As Variant
    Dim StrDataType As String
    Dim i As Integer
    Dim DataArray(1 To 48) As Variant
    
    StationID = "'" & ModGlobals.ConvStnNotoID([StationNo]) & "'"
    
    Select Case DataType
        Case Is = PencilGraphType.NoCrew
            StrDataType = "CrewNo"
        Case Is = PencilGraphType.NoDrivers
            StrDataType = "DriverNum"
        Case Is = PencilGraphType.NoJOs
            StrDataType = "JONum"
    End Select
    
    Set RstPencilData = DBase.SQLQuery("SELECT * FROM stationdetail WHERE tid = " & MonthNo & _
                                        "AND dateslot = " & DateNo & _
                                        "AND stationid = " & StationID & _
                                        "ORDER BY timeslot ASC")
    
    With RstPencilData
        For i = 1 To 48
            If Not .EOF Then
                If !Timeslot = i Then
                    DataArray(i) = .Fields(StrDataType)
                    PrevData = !CrewNo
                    .MoveNext
                Else
                DataArray(i) = PrevData
                End If
            Else
                DataArray(i) = PrevData
            End If
        Next
    End With
    
    GetData = DataArray()
    
    Set RstPencilData = Nothing

End Function

Private Sub Class_Initialize()
    Set pShapes = New Collection
End Sub

Private Sub Class_Terminate()
    Set pShapes = Nothing
End Sub
