Attribute VB_Name = "Main"
Option Explicit


Public Sub Initialise()
    Set DBase = New ClsDatabase
    Library.PerfSettingsOff
    Summary.InitialiseArrows
    Application.Worksheets(1).Activate
End Sub


Public Sub CloseDown()
    Set DBase = Nothing
End Sub
