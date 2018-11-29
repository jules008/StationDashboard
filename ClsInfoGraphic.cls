VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsInfoGraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim pEfficiency As Single
Dim pAvailability As Single
Dim pTarget As Single
Dim pStation As Integer

'Efficiency property
'------------------
Public Property Let Efficiency(Value As Single)
    pEfficiency = Value
End Property

'Availability property
'------------------
Public Property Let Availability(Value As Single)
    pAvailability = Value
End Property

'Target property
'------------------
Public Property Let Target(Value As Single)
    pTarget = Value
End Property

'Station property
'------------------
Public Property Let Station(Value As Integer)
    pStation = Value
End Property

Public Sub UpdateGraph(Station As Integer, Efficiency As Single, Availability As Single, Target As Single)
    Dim ActiveShape As Shape
    
    Set ActiveShape = ActiveWorkbook.Worksheets("Summary").Shapes(Station & "_Efficiency")

    ActiveShape.Width = Efficiency
    
    Set ActiveShape = ActiveWorkbook.Worksheets("Summary").Shapes(Station & "_Availability")

    ActiveShape.Width = Availability
    
'    Set ActiveShape = ActiveWorkbook.Worksheets("Summary").Shapes(Station & "_Target")

'    ActiveShape.Left = 377.6 + Target
    
End Sub


Public Sub SelectRectangle()
    Dim MyShape As Shape
    
    For Each MyShape In ActiveSheet.Shapes
    
        If MyShape.AutoShapeType = msoShapeRoundedRectangle Then
            MyShape.Select Replace:=False
    
        End If
    Next
End Sub
Public Sub SelectHorizLine()
    Dim MyShape As Shape
    
    For Each MyShape In ActiveSheet.Shapes
    
        If MyShape.Height = 0 Then
            MyShape.Select Replace:=False
    
        End If
    Next
End Sub
Public Sub SelectVertLine()
    Dim MyShape As Shape
    
    For Each MyShape In ActiveSheet.Shapes
    
        If MyShape.Width = 0 Then
            MyShape.Select Replace:=False
    
        End If
    Next
End Sub

