VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 21 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' Show - makes arrow visible
' Hide - makes arrow invisible
' Toggle - Toggles arrow direction
'===============================================================
Option Explicit
Private pArrow As Shape
Private pDir As Direction
Private pDataItem As DataCol

'------------------------------------------------------------------------------
Public Property Get Arrow() As Shape
    Set Arrow = pArrow
End Property

Public Property Let Arrow(ByVal vNewValue As Shape)
    Set pArrow = vNewValue
End Property

'------------------------------------------------------------------------------
Public Property Get Dir() As Direction
    Dir = pDir
End Property

Public Property Let Dir(ByVal vNewValue As Direction)
    pDir = vNewValue
End Property

'------------------------------------------------------------------------------
Public Property Get DataItem() As DataCol
    DataItem = pDataItem
End Property

Public Property Let DataItem(ByVal vNewValue As DataCol)
    pDataItem = vNewValue
End Property

'===============================================================
'Method Show
'makes arrow visible
'---------------------------------------------------------------
Public Sub Show()
    pArrow.Visible = msoCTrue
End Sub

'===============================================================
'Method Hide
' makes arrow invisible
'---------------------------------------------------------------
Public Sub Hide()
    pArrow.Visible = msoFalse
End Sub

'===============================================================
' Method Toggle
' rotates arrow Toggle
'---------------------------------------------------------------
Public Sub Toggle()
    If pDir = DirUp Then
        pArrow.Rotation = 90
        pDir = DirDown
    Else
        pArrow.Rotation = 270
        pDir = DirUp
    End If
End Sub





