VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub Workbook_Open()
    Initialise
    MsgBox "Please note that the methodology used in this report differs from that of the current standard methodology used for reporting across the service, and on the divisional tool. Use this dashboard for reference and station management only, but not for final reporting figures.", vbOKOnly + vbInformation, "Station Dashboard"
End Sub
