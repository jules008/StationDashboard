VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPencilGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 26 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds training item from index
' AddItem - Add training sesssion to collection
'===============================================================
Option Explicit

Private pPencilGraphics As Collection

'===============================================================
'Method FindItem
'finds training item in collection
'---------------------------------------------------------------
Public Function FindItem(PencilGraphic As Variant) As ClsPencilGraphic
    On Error Resume Next
    Set FindItem = pPencilGraphics.Item(PencilGraphic)
End Function

'===============================================================
'Method AddItem
'Add training sesssion to collection
'---------------------------------------------------------------
Public Sub AddItem(PencilGraphic As ClsPencilGraphic, Optional IndexKey As String)
    pPencilGraphics.Add PencilGraphic
End Sub

'===============================================================
'Method removeItem
'deletes training item from index
'---------------------------------------------------------------
Public Sub RemoveItem(Index As Variant)
    pPencilGraphics.RemoveItem Index
End Sub

'===============================================================
'Method count
'returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pPencilGraphics.Count
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pPencilGraphics = New Collection
End Sub

Private Sub Class_Terminate()
    Set pPencilGraphics = Nothing
End Sub
'===============================================================



