VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsArrows"
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
' AddItem - add item to collection
' HideAll - hides all arrows in the collection
'===============================================================
Option Explicit
Private pArrows As Collection

'===============================================================
'Method HideAll
'hides all arrows in the collection
'---------------------------------------------------------------
Public Sub HideAll()
    Dim Obj As ClsArrow
    For Each Obj In pArrows
        Obj.Hide
    Next
End Sub

'===============================================================
'Method AddItem
'Add Item to collection
'---------------------------------------------------------------
Public Sub AddItem(Arrow As ClsArrow)
    pArrows.Add Arrow, key:=CStr(Arrow.DataItem)
End Sub

'===============================================================
Private Sub Class_Initialize()
    Set pArrows = New Collection
End Sub

Private Sub Class_Terminate()
    Set pArrows = Nothing
End Sub



