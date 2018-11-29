Attribute VB_Name = "ModMain"
Option Explicit


Public Sub Initialise()
    Set DBase = New ClsDatabase
    ModLibrary.PerfSettingsOff
    ShtSummary.InitialiseArrows
    Application.Worksheets(1).Activate
End Sub


Public Sub CloseDown()
    Set DBase = Nothing
End Sub
