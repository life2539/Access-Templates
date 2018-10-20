Attribute VB_Name = "refAccessFunctions"
Option Compare Database
Function RenameFieldName(TargetTable As String, TargetField As String, NewFieldName As String, Optional DBfilename As String)

    Dim dbs As DAO.database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    db = OpenDatabase("DB Name")
    If DBfilename = "" Then
        Set dbs = CurrentDb
    Else
        Set dbs = OpenDatabase(DBfilename)
    End If
    
    Set tdf = dbs.TableDefs(TargetTable)
    Set fld = tdf.Fields(TargetField)
    fld.Name = NewFieldName
    
    dbs.Close
    Set dbs = Nothing
    Set fld = Nothing
    Set tdf = Nothing
End Function
Function DoSQL(SQLString As String, Optional ShowWarnings As Boolean = False, Optional RecordActions As Boolean = True, Optional LogMessage As String = "DOSQL... ") As Boolean
    On Error GoTo Errorhandler
    With DoCmd
        .SetWarnings (ShowWarnings)
        .RunSQL SQLString
        DoSQL = True
        If RecordActions Then LogRecord (LogMessage & "Success")
        .SetWarnings True
    End With
    Exit Function
Errorhandler:
        DoCmd.SetWarnings True
        
        DoSQL = False
        If Verbose Then MsgBox "SQL Error " & Err & " " & Error & ". Check VBA Immediate window Dialog. - Alex"
        If RecordActions Then LogRecord ("SQL Error " & Err & " " & Error)
        If RecordActions Then LogRecord (SQLString)
        
End Function
Public Function TableExists(Tablename As String)
    TableExists = ObjectExists(Tablename, acTable)
End Function
Public Function QueryExists(QueryName As String)
    QueryExists = ObjectExists(QueryName, "QUERY")
End Function


Public Function ObjectExists(ObjectName As String, Optional ObjectType As AcObjectType = acDefault)
    Select Case ObjectType
    Case acTable
        ObjectExists = Not IsNull(DLookup("Name", "MSysObjects", "Name='" & ObjectName & "' And Type In (1,4,6)"))
    Case acQuery
        ObjectExists = Not IsNull(DLookup("Name", "MSysObjects", "Name='" & ObjectName & "' And Type In (5)"))
    Case acModule
        ObjectExists = Not IsNull(DLookup("Name", "MSysObjects", "Name='" & ObjectName & "' And Type In (-32761)"))
    Case acDefault
        ObjectExists = Not IsNull(DLookup("Name", "MSysObjects", "Name='" & ObjectName & "'"))
    End Select
End Function
Public Function LogRecord(Optional Msg As String = "New Record")
    DoCmd.SetWarnings False
    If Not TableExists("RecordLog") Then
        
        'DoCmd.RunSQL "SELECT Now() AS [TimeStamp], '" & Msg & "' AS [Message] INTO RecordLog;"  '"CREATE TABLE [RecordLog] ([TimeStamp] DATE, [Message] MEMO)"
    Else
        'DoCmd.RunSQL "INSERT INTO RecordLog ( [TimeStamp], Message ) SELECT Now() AS TimeStamp, '" & Msg & "' AS Message;"
    End If
    LogRecord = Msg
    DoCmd.SetWarnings True
End Function


Function Eomonth(DateValue As Date, Optional Months As Integer = 0) As Date
    Eomonth = DateSerial(Year(DateValue), Month(DateValue) + Months + 1, 0)
End Function
Function Somonth(DateValue As Date, Optional Months As Integer = 0) As Date
    Somonth = Eomonth(DateValue, Months) + 1
End Function

