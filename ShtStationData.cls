VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtStationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub GetAvailability(Month As Integer, Data_2 As Boolean, Data_1 As Boolean, Data As Boolean, Data1 As Boolean, Order As String)

    Dim TableName As String
    Dim FieldName As String
    Dim TargetRange As Range
    Dim RecSet As Recordset
    Dim SQL As String
    Dim DataSheet As Worksheet
    Dim Data_2Range As Range
    Dim Data_1Range As Range
    Dim DataRange As Range
    Dim Data1Range As Range
    
    'set Range objects
    Set Data_2Range = Range("Data_Range_2")
    Set Data_1Range = Range("Data_Range_1")
    Set DataRange = Range("Data_Range")
    Set Data1Range = Range("Data_Range1")
    
    'Get selected months data
    '------------------------
    If Data = True Then
    
        'set target range to where the data will be written
        Set TargetRange = DataRange
        
        'run SQL query
        SQL = "SELECT * FROM Station WHERE tid= " & Month & " ORDER BY " & Order
'        Debug.Print SQL
        
        Set RecSet = DBase.SQLQuery(SQL)
        
        'write values
        TargetRange.CopyFromRecordset RecSet
    End If
    
    'Get month - 1 data
    '------------------------
    If Data_1 = True Then
        'set target range to where the data will be written
        Set TargetRange = Data_1Range
        
        'run SQL query
        SQL = "SELECT * FROM Station WHERE tid= " & Month - 1 & " ORDER BY " & Order
        Set RecSet = DBase.SQLQuery(SQL)
        
        'write values
        TargetRange.CopyFromRecordset RecSet
    End If
    
    'Get month - 2 data
    '------------------------
    If Data_2 = True Then
        'set target range to where the data will be written
        Set TargetRange = Data_2Range
        
        'run SQL query
        SQL = "SELECT * FROM Station WHERE tid= " & Month - 2 & " ORDER BY " & Order
        Set RecSet = DBase.SQLQuery(SQL)
        
        'write values
        TargetRange.CopyFromRecordset RecSet
        
    End If
    
    'Get month + 1 data
    '------------------------
    If Data1 = True Then
        'set target range to where the data will be written
        Set TargetRange = Data1Range
        
        'run SQL query
        SQL = "SELECT * FROM Station WHERE tid= " & Month + 1 & " ORDER BY " & Order
        Set RecSet = DBase.SQLQuery(SQL)
        
        'write values
        TargetRange.CopyFromRecordset RecSet
    End If
    
    'update month string
    '----------------------
    'get Date string from month integer value
    Set TargetRange = Worksheets("Summary").Range("MonthString")
    
    'run SQL
    SQL = "SELECT MonthStamp FROM TimeTbl WHERE tid = " & Month
    Set RecSet = DBase.SQLQuery(SQL)
    
    TargetRange.CopyFromRecordset RecSet
    
    
    'get number of days in month
    Summary.Range("NoDays") = ModLibrary.dhDaysInMonth(Summary.Range("U5"))
    
    'close down
    Set RecSet = Nothing
    Set Data_2Range = Nothing
    Set Data_1Range = Nothing
    Set DataRange = Nothing
    Set Data1Range = Nothing


End Sub
Public Function NextMonth(MonthNum As Integer, MinVal As Integer, MaxVal As Integer) As Integer
    Dim Data_2Range As Range
    Dim Data_1Range As Range
    Dim DataRange As Range
    Dim Data1Range As Range
    
    'set Range objects
    Set Data_2Range = Range("Data_Range_2")
    Set Data_1Range = Range("Data_Range_1")
    Set DataRange = Range("Data_Range")
    Set Data1Range = Range("Data_Range1")
        
    Select Case MonthNum
        Case Is = MaxVal
            'at top of list, so do nothing
            
        Case Is = MaxVal - 1
            'increment month no
            MonthNum = MonthNum + 1
             
            'shift data by a month
            Data_2Range.Value = Data_1Range.Value
            Data_1Range.Value = DataRange.Value
            DataRange.Value = Data1Range.Value
            
            'clear data - 1 range as there is no data
            Data1Range.ClearContents
            
            'get remaining data
            GetAvailability MonthNum, False, False, False, False, Summary.Range("orderBy")
                 
        Case Else
            'increment month no
            MonthNum = MonthNum + 1
            
            'shift data by a month
            Data_2Range.Value = Data_1Range.Value
            Data_1Range.Value = DataRange.Value
            DataRange.Value = Data1Range.Value
            
            'get remaining data
            GetAvailability MonthNum, False, False, False, True, Summary.Range("orderBy")
            
    End Select
       
    'update month no on sheet
    NextMonth = MonthNum

End Function

Public Function PrevMonth(MonthNum As Integer, MinVal As Integer, MaxVal As Integer) As Integer
    Dim Data_2Range As Range
    Dim Data_1Range As Range
    Dim DataRange As Range
    Dim Data1Range As Range
    
    'set Range objects
    Set Data_2Range = Range("Data_Range_2")
    Set Data_1Range = Range("Data_Range_1")
    Set DataRange = Range("Data_Range")
    Set Data1Range = Range("Data_Range1")
        
    Select Case MonthNum
        Case Is = MinVal
            'at bottom of list, so do nothing
            
        Case Is = MinVal + 1
            'decrement month no
            MonthNum = MonthNum - 1
             
            'shift data by a month
            Data1Range.Value = DataRange.Value
            DataRange.Value = Data_1Range.Value
            Data_1Range.Value = Data_2Range.Value
            
            'clear data - 1 range as there is no data
            Data_1Range.ClearContents
           
            'get remaining data
            GetAvailability MonthNum, False, False, True, False, Summary.Range("orderBy")
        
        Case Is = MinVal + 2
            'decrement month no
            MonthNum = MonthNum - 1
            
            'shift data by a month
            Data1Range.Value = DataRange.Value
            DataRange.Value = Data_1Range.Value
            Data_1Range.Value = Data_2Range.Value
            
            'clear data - 2 range as there is no data
            Data_2Range.ClearContents
            
            'get remaining data
            GetAvailability MonthNum, False, False, False, False, Summary.Range("orderBy")
     
        Case Else
            'decrement month no
            MonthNum = MonthNum - 1
            
            'shift data by a month
            Data1Range.Value = DataRange.Value
            DataRange.Value = Data_1Range.Value
            Data_1Range.Value = Data_2Range.Value
            
            'get remaining data
            GetAvailability MonthNum, True, False, False, False, Summary.Range("orderBy")
    End Select
    PrevMonth = MonthNum
End Function




