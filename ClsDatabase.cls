VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - added error handling for database not found
'---------------------------------------------------------------
' Date - 28 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' SQLQuery - Query database
' DBConnect - make Database connection
'===============================================================

Option Explicit
Private pDBPath As String
Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef

'---------------------------------------------------------------
Public Property Get DBPath() As String
    DBPath = pDBPath
End Property

'===============================================================
' Method SQLQuery
' Query database
'---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim Results As Recordset
    
    Set SQLQuery = DB.OpenRecordset(SQL, dbOpenDynaset)
    
End Function

'===============================================================
' Method DBConnect
' make database connection
'---------------------------------------------------------------
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
End Sub

'---------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo errorhandler
    
    Dim wrkspc As DAO.Workspace
    Set wrkspc = DBEngine(0)
    pDBPath = Range("DBpath")
    Set DB = wrkspc.OpenDatabase(DBPath, False, False)

errorhandler:
    If Err.Number = 3044 Then
        MsgBox "Database not found, please update location"
        DBConnect
        ModMain.Initialise
    End If
End Sub

Private Sub Class_Terminate()
    Set DB = Nothing
End Sub
'---------------------------------------------------------------

