Attribute VB_Name = "ModLibrary"
Option Explicit

Public Function ConvertHoursIntoDecimal(TimeIn As Date)
    Dim TB, Result As Single
    
    TB = Split(TimeIn, ":")
    ConvertHoursIntoDecimal = TB(0) + ((TB(1) * 100) / 60) / 100
    
End Function
Function EndOfMonth(InputDate As Date) As Variant
    EndOfMonth = Day(DateSerial(Year(InputDate), Month(InputDate) + 1, 0))
End Function
Public Sub DBConnect()
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly, "No Files"

        Exit Sub
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    'save database location
    Range("dbpath") = FileLoc
    Set DlgOpen = Nothing

End Sub
Public Sub PerfSettingsOn()

    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

Public Sub PerfSettingsOff()
    
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
Function dhDaysInMonth(Optional dtmDate As Date = 0) As Integer
    ' Return the number of days in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhDaysInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 1) - _
     DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function

Public Function ConvStnNotoID(StationNo As Integer) As String
    ConvStnNotoID = "EC" & CStr(Format(StationNo, "00"))

End Function
