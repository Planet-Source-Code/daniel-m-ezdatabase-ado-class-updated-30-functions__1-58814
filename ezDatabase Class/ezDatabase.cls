VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ezDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author Comments/Information :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
'This class was written by Daniel M. Use it freely in your projects as you like.        '
'It is commented but could be better. Example application shows more. This is my first  '
'class so please give me comments at seoulxkorean@yahoo.com or at planet-source-code.com'
'Thanks for downloading!                                                                '
':::::::::::::::::::::::::::::::::::::::::::::::::: Wow I love these comment boxes =D   '

'Please give credit if you do intend on using this in your projects. Thanks.
Option Explicit

'Database connection variables, these must be referenced.
Private dbCon           As New ADODB.Connection
Private dbRs()          As New ADODB.Recordset

'Query variables
Private bFindAll        As Boolean  'Is FindAll being used?
Private bLoop           As Boolean  'Still searching for query?
Private rsIndex         As Long     'Current index being queried.
Private rsLastField     As String   'Last Field searched.
Private rsLastQuery     As String   'Last Query.
Private typeCriteria    As String   'Criteria string.

'Property Variables
Private bMode           As Byte     'Connection Mode
Private bWritable       As Boolean  'Is MDB File Writable?
Private bBackup         As Boolean  'Create Backup?
Private bCaseSensitive  As Boolean  'Is Case Sensitive?
Private bConnected      As Boolean  'Is there a connection?
Private bFindMarker     As Boolean  'Was Query found?
Private sConnection     As String   'Database Location String
Private sUsername       As String   'Username String
Private sPassword       As String   'Password String
Private rsRequest       As Long     'Request Type
Private tblCount        As Long     'Table Count

    
'Table Connection String Collection
Private tblString       As New Collection

'Request Type Enumeration: Based on request type, function will only work inside those parameters.
Enum RequestType
    rtCurrentTable = 1 'Function works only within current table/recordset to retrieve data
    rtAllTables = 2    'Function works with all tables/recordsets to retrieve data
End Enum




'Error Message Constants
Private Const errDatabaseUndefined      As String = "Database Connection not Defined."
Private Const errTableUndefined         As String = "No tables have been defined."
Private Const errNotConnected           As String = "Must be Connected to Query Data."
Private Const errActionWhileConnected   As String = "Action cannot be performed while connected."
Private Const errTableNotConnected      As String = "Could not retrieve data from table."
Private Const errStartEndofFile         As String = "Record is either BOF or EOF, Action Canceled."

'Public Events
Public Event InvalidParameter(ByRef Description As String)  'Specific Error-Type Event
Public Event ConnectionError(ByRef Description As String)   'Connection Error-Type Event
Public Event Connect()                                      'On Connect
Public Event Disconnect()                                   'On Disconnect
Public Event FindNotUsed()                                  'Specific Error-Type Event
Public Event FieldsChanged(ByRef QueryInfo As String)       'Backbone Event used to retrieve data
Public Event Error(ByRef Description As String)             'All miscellaneous errors go here.
Public Event QueryNotFound(ByRef Query As String)           'Backbone event for unfound data.
Public Event Added()                                        'Successful Event
Public Event Deleted()                                      'Successful Event
Public Event Updated()                                      'Successful Event
'========================= Initialize/Connection & Disconnection Subs =========================
Private Sub Class_Initialize()
    ConnectMode = adModeReadWrite 'Default Property
End Sub

Public Sub Expose()
If Connected Then
    RaiseEvent ConnectionError("Already Connected")
Else
    If Len(dbLocation) Then
        If tblCount > 0 Then
            dbCon.Mode = Mode
            dbCon.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
            "data source=" & dbLocation & ";" 'Connection information
            If Backup Then
                If Dir(dbLocation) <> "" Then
                    Dim sSplit() As String, sPath As String
                        sSplit$ = Split(dbLocation$, "\")
                        sPath$ = Left$(dbLocation$, Len(dbLocation$) - Len(sSplit(UBound(sSplit$)))) & "Backup"
                        If FolderExists(sPath) = False Then
                            MkDir (sPath$) 'Make Backup folder if non-existant
                        End If
                        FileCopy dbLocation$, sPath$ & "\" & sSplit(UBound(sSplit$))
                        'Will overwrite if any previous existing backups.
                End If
            End If
            checkAttrib 'Check if .mdb is writable
            dbCon.CursorLocation = adUseClient
            dbCon.Open dbCon.ConnectionString, Username$, Password$ 'Open the connection
            
            If dbCon.State = 1 Then
                Dim i As Long
                For i = 0 To tblCount - 1
                    If dbRs(i).State = 0 Then 'Hopefully it is at this state
                        dbRs(i).Open "Select * From [" & tblString(i + 1) & "]", dbCon, adOpenDynamic, adLockOptimistic
                    Else
                        dbRs(i).Close 'If already opened, close it then re-open
                        dbRs(i).Open "Select * From [" & tblString(i + 1) & "]", dbCon, adOpenDynamic, adLockOptimistic
                    End If
                Next i
                RaiseEvent Connect
                bConnected = True 'Update status
            Else
                RaiseEvent ConnectionError("Could not connect to '" & sConnection$ & "'")
            End If
        Else
            RaiseEvent ConnectionError(errTableUndefined$)
        End If
    Else
        RaiseEvent ConnectionError(errDatabaseUndefined$)
    End If
End If
End Sub
Public Sub Dispose()
If Connected = True Then
    RaiseEvent Disconnect
    bConnected = False
    
    Dim i As Long
        For i = 0 To UBound(dbRs) 'Prevent errors by checking their states before closing
            If dbRs(i).State = 1 Then dbRs(i).Close 'shouldn't happen, but just in case.
            Set dbRs(i) = Nothing 'Clear from memory
        Next i
        
        If dbCon.State = 1 Then dbCon.Close
        Set dbCon = Nothing 'Clear from memory
        tblCount = 0
        Do Until tblString.Count = 0 'Clear collection
            tblString.Remove 1
        Loop
Else
    RaiseEvent Error("You are already Disconnected.") 'Can't dispose whats not already there!
End If
End Sub
'=========================  Let Property Values  =====================================
Public Property Let ConnectMode(Mode As ConnectModeEnum)
bMode = Mode
End Property
Public Property Let dbLocation(ByRef mdbLocation As String)
    If Dir(mdbLocation) <> "" Then 'Method to test if a file exists
        sConnection$ = mdbLocation$ 'Set location of .mdb file
    Else
        RaiseEvent InvalidParameter("Database File does not exist.")
    End If
End Property

Public Property Let Username(ByRef Username As String)
    sUsername$ = Username$ 'Set username for database
End Property
Public Property Let Password(ByRef Password As String)
    sPassword$ = Password$ 'Set password for database
End Property
Public Property Let RequestType(ByVal RequestType As RequestType)
    rsRequest = RequestType
End Property
Public Property Let CaseSensitive(ByVal CaseSensitive As Boolean)
bCaseSensitive = CaseSensitive
End Property
Public Property Let Backup(ByRef Backup As Boolean)
bBackup = Backup
End Property
'========================  Get Property Values  ======================================
Public Property Get Mode() As Byte
Mode = bMode
End Property
Public Property Get Backup() As Boolean
Backup = bBackup
End Property
Public Property Get CaseSensitive() As Boolean
CaseSensitive = bCaseSensitive
End Property
Public Property Get RequestType() As Long
RequestType = rsRequest
End Property
Public Property Get FindMarker() As Boolean
FindMarker = bFindMarker
End Property
Public Property Get RecordCount(Optional ByVal CountType As RequestType = 2) As Long
If Connected Then
    If CountType = 1 Then   'Get record count for current table/recordset
        RecordCount = dbRs(rsIndex).RecordCount
    Else
        Dim i As Long
        For i = 0 To UBound(dbRs) 'Get record count for all tables/recordsets
            RecordCount = RecordCount + dbRs(i).RecordCount
        Next i
    End If
Else
    RaiseEvent Error(errNotConnected$)
End If
End Property
Public Property Get Username() As String
Username$ = sUsername$ 'Get username specified by user
End Property
Public Property Get Password() As String
Password$ = sPassword$ 'Get password specified by user
End Property
Public Property Get BOF(Optional ByVal RequestType As RequestType = 2) As Boolean
If Connected Then
    If RequestType <> 1 Then
        If rsIndex = 0 Then
            BOF = dbRs(rsIndex).BOF
        Else
            BOF = False
        End If
    Else
        BOF = dbRs(rsIndex).BOF
    End If
Else
    RaiseEvent Error(errNotConnected$)
End If
End Property
Public Property Get EOF(Optional ByVal RequestType As RequestType = 2) As Boolean
If Connected Then
    If RequestType <> 1 Then
        If rsIndex = tblCount - 1 Then EOF = dbRs(rsIndex).EOF Else EOF = False 'Similar to BOF
    Else
        EOF = dbRs(rsIndex).EOF
    End If
Else
    RaiseEvent Error(errNotConnected$)
End If
End Property
Public Property Get Field(ByRef NameOrIndex As Variant) As String
If Connected Then
    If dbRs(rsIndex).State = 1 Then 'Make sure we're connected
        If dbRs(rsIndex).BOF Or dbRs(rsIndex).EOF Then 'If BOF/EOF then retrieved data is ""
            Field$ = ""
        Else
            If IsNull(dbRs(rsIndex)(NameOrIndex)) = True Then 'If field is null, we cannot give null
                Field$ = ""                                 'data so we must handle this.
            Else
                Field$ = dbRs(rsIndex)(NameOrIndex)
            End If
        End If
    Else
        RaiseEvent ConnectionError(errTableNotConnected$)
    End If
Else
    RaiseEvent Error(errNotConnected$)
End If
End Property
Public Property Get Connected() As Boolean
    Connected = bConnected 'Get Connected State
End Property
Public Property Get TableCount() As Long
    TableCount = tblCount 'Get table count
End Property
Public Property Get dbLocation() As String
    dbLocation$ = sConnection$ 'Get location to .mdb
End Property
Public Property Get CurrentTable() As String
    If Connected Then CurrentTable$ = tblString.Item(rsIndex + 1) Else CurrentTable = ""
End Property
Public Property Get TableName(ByVal TableNum As Long) As String
If tblString.Count Then
    If ValidInt(TableNum, 0, tblString.Count) Then
        TableName$ = tblString.Item(TableNum + 1) 'Give tablename of specified tablenum
    Else
        RaiseEvent Error("Invalid Table Number.") 'Error handling
    End If
Else
    RaiseEvent Error("No Tables have been added.")
End If
End Property
Public Property Get TableNumber() As Long
    TableNumber = rsIndex 'Get current table number
End Property
Public Property Get LastQuery() As String
    LastQuery$ = rsLastQuery$ 'Get last query
End Property
'=============================     Class Functions / Subs     ===============================
Public Sub AddTable(ByRef NameOfTable As String)
If Connected Then 'Error handling once again.
    RaiseEvent Error(errActionWhileConnected$)
Else
    If Len(NameOfTable$) Then
        If tblCount = 0 Then 'If first table then redim as 0
            ReDim dbRs(0)
        Else 'otherwise increase count
            ReDim dbRs(UBound(dbRs) + 1)
        End If
        tblString.Add NameOfTable$ 'add table name
        tblCount = tblCount + 1 'add table count
    Else
        RaiseEvent InvalidParameter("Table Name must be at least one character length.")
    End If
End If

End Sub
Public Sub RemoveTable(ByVal TableNum As Long)
If Connected Then
    RaiseEvent Error(errActionWhileConnected$)
Else
    If tblString.Count = 0 Then
        RaiseEvent InvalidParameter("Table Number is Invalid")
    Else
        tblString.Remove TableNum
        tblCount = tblCount - 1
    End If
End If
End Sub
Public Sub UpdateTable(ByVal TableNum As Long, ByRef NewTableName As String)
If Connected Then
    RaiseEvent Error(errActionWhileConnected$)
Else
    If tblString.Count = 0 Or TableNum > tblString.Count Or TableNum < 1 Or NewTableName = "" Then
        RaiseEvent InvalidParameter("Table Number/Name is Invalid")
    Else
        tblString.Item(TableNum) = NewTableName$ 'update table name by tablenum
    End If
End If
End Sub

Public Sub Update(ByRef rsField As String, ByRef rsNewValue As String)
If Connected Then
    If bWritable = False Then RaiseEvent Error("Database File is set to Read Only."): Exit Sub
    If dbRs(rsIndex).State = 1 Then
        If dbRs(rsIndex).BOF Or dbRs(rsIndex).EOF Then
            RaiseEvent Error(errStartEndofFile$)
        Else
            dbRs(rsIndex)(rsField) = rsNewValue$
            dbRs(rsIndex).Update
            RaiseEvent Updated
            RaiseEvent FieldsChanged("Record Updated")
        End If
    Else
        RaiseEvent Error(errTableNotConnected$)
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If

End Sub

Public Sub Delete()
If Connected Then
    If bWritable = False Then RaiseEvent Error("Database File is set to Read Only."): Exit Sub
    If dbRs(rsIndex).BOF Or dbRs(rsIndex).EOF Then
        RaiseEvent Error(errStartEndofFile$)
    Else
        dbRs(rsIndex).Delete
        dbRs(rsIndex).MovePrevious
        RaiseEvent Deleted
        RaiseEvent FieldsChanged("Record Deleted")
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub

Public Sub AddNew(ByVal tblNumber As Long, ByRef objArray As Variant, Optional ByVal iStart As Long = 0)
If Connected Then
    If bWritable = False Then RaiseEvent Error("Database File is set to Read Only."): Exit Sub
    If tblNumber > tblCount - 1 Or tblNumber < 0 Then
        RaiseEvent InvalidParameter("Table Number Invalid")
    Else
        rsIndex = tblNumber 'make current table the given value
        Dim i As Long, f As Long, iEnd As Long
        f = 1: i = iStart
        
        If UBound(objArray) - iStart < dbRs(rsIndex).Fields.Count Then
            RaiseEvent InvalidParameter("Record could not be updated due to invalid Array.")
        Else
            With dbRs(rsIndex)
                .MoveLast 'move last
                .AddNew 'add new
                    Do
                        .Fields((f)) = objArray(i)
                        f = f + 1
                        i = i + 1
                    Loop Until f = .Fields.Count
                .Update 'update the fields
            End With
            RaiseEvent Added
            RaiseEvent FieldsChanged("Record Added")
        End If
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub MoveRecordset(ByVal TableNum As Long)
If Connected Then
    If ValidInt(TableNum, 0, UBound(dbRs)) = True Then
        rsIndex = TableNum
        dbRs(rsIndex).MoveFirst
        RaiseEvent FieldsChanged("MoveRecordset:" & TableNum)
    Else
        RaiseEvent InvalidParameter("Invalid Table Number")
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub FirstRecordset()
If Connected Then
    rsIndex = 0
    dbRs(rsIndex).MoveFirst
    RaiseEvent FieldsChanged("MoveFirstRecordset:" & rsIndex)
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub NextRecordset()
If Connected Then
    If rsIndex <> tblCount - 1 Then rsIndex = rsIndex + 1 Else Exit Sub
    dbRs(rsIndex).MoveFirst
    RaiseEvent FieldsChanged("MoveNextRecordset:" & rsIndex)
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub PreviousRecordset()
If Connected Then
    If rsIndex <> 0 Then rsIndex = rsIndex - 1 Else Exit Sub
    dbRs(rsIndex).MoveFirst
    RaiseEvent FieldsChanged("MovePreviousRecordset:" & rsIndex)
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub LastRecordset()
If Connected Then
    rsIndex = tblCount - 1
    dbRs(rsIndex).MoveFirst
    RaiseEvent FieldsChanged("MoveLastRecordset:" & rsIndex)
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub Move(ByVal NumRecords As Long, Optional ByVal Start As Long = 0)
If Connected Then
    If dbRs(rsIndex).EOF Then
        If rsIndex <> tblCount - 1 Then rsIndex = rsIndex + 1 Else rsIndex = 0
        dbRs(rsIndex).MoveFirst
    End If
    
    If Start <> 0 Then
        dbRs(rsIndex).Move NumRecords, Start
        RaiseEvent FieldsChanged("Move:" & NumRecords & "," & Start)
    Else
        dbRs(rsIndex).Move NumRecords
        RaiseEvent FieldsChanged("Move:" & NumRecords)
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If

End Sub
Public Sub MoveFirst(Optional ByVal RequestType As RequestType = 2)
If Connected Then
    If RequestType <> 1 Then rsIndex = 0 'set recordset index to 0
    dbRs(rsIndex).MoveFirst 'move first
    RaiseEvent FieldsChanged("MoveFirst:" & RequestType)
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub MoveNext()
If Connected Then
    With dbRs(rsIndex)
        If .EOF Then 'end of file
            If rsIndex = tblCount - 1 Then
                'do nothing
            Else
                rsIndex = rsIndex + 1
                dbRs(rsIndex).MoveFirst
                RaiseEvent FieldsChanged("MoveNext")
            End If
        Else
        .MoveNext
            If .EOF Then
                If rsIndex = tblCount - 1 Then
                    'do nothing
                Else
                    rsIndex = rsIndex + 1
                    dbRs(rsIndex).MoveFirst
                    RaiseEvent FieldsChanged("MoveNext")
                End If
            Else
                RaiseEvent FieldsChanged("MoveNext")
            End If
        End If
    End With
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub MovePrev()
If Connected Then
    With dbRs(rsIndex)
        If .BOF Then
            If rsIndex = 0 Then
                'do nothing
            Else
                rsIndex = rsIndex - 1
                dbRs(rsIndex).MoveLast
                RaiseEvent FieldsChanged("MovePrev")
            End If
        Else
        .MovePrevious
            If .BOF Then
                If rsIndex = 0 Then
                    'do nothing
                Else
                    rsIndex = rsIndex - 1
                    dbRs(rsIndex).MoveLast
                    RaiseEvent FieldsChanged("MovePrev")
                End If
            Else
                RaiseEvent FieldsChanged("MovePrev")
            End If
        End If
    End With
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub MoveLast(Optional ByVal RequestType As RequestType = 2)
If Connected Then
    If RequestType <> 1 Then rsIndex = tblCount - 1
    dbRs(rsIndex).MoveLast
    RaiseEvent FieldsChanged("MoveLast")
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub Find(ByRef nField As String, ByRef sQuery As String, Optional ByVal RequestType As RequestType = 2)
FindFirst nField$, sQuery$, RequestType
End Sub
Public Sub FindAll(ByRef nField As String, ByRef sQuery As String, Optional ByVal RequestType As RequestType = 2)
If Connected Then 'FindAll uses findfirst then findnext until end of file on all tables

    If setCriteria(nField$, sQuery$) = False Then Exit Sub
    bFindMarker = False
    FindFirst nField$, sQuery$, RequestType
    
    
    bFindAll = True
    If EOF(2) = False Then
        Do Until EOF(2) Or bFindAll = False
            FindNext
        Loop
    End If
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If

End Sub
Public Sub FindFirst(ByRef nField As String, ByRef sQuery As String, Optional ByVal RequestType As RequestType = 2)
If Connected Then
    If setCriteria(nField$, sQuery$) = False Then Exit Sub
    bFindMarker = False
    rsRequest = RequestType
    rsIndex = 0
    
    Do
        dbRs(rsIndex).MoveFirst
        dbRs(rsIndex).Find typeCriteria$, 0, adSearchForward
                
        If dbRs(rsIndex).EOF Then
            If RequestType <> 1 Then
                If rsIndex <> tblCount - 1 Then
                    rsIndex = rsIndex + 1
                Else
                    RaiseEvent QueryNotFound(sQuery$) 'Query not found!
                    rsLastQuery$ = vbNullString
                    bLoop = True
                End If
            Else
                RaiseEvent QueryNotFound(sQuery$)
                rsLastQuery$ = vbNullString
                bLoop = True
            End If
        Else
            If CaseSensitive = True Then
                Dim sResult As String
                sResult = dbRs(rsIndex).Fields(nField)
                    If Trim(sQuery) = Trim(sResult) Then
                        RaiseEvent FieldsChanged("FindFirst:" & nField$) 'Query found
                        bFindMarker = True
                        bLoop = True
                    Else
                        FindNext
                        Exit Sub
                    End If
            Else
                RaiseEvent FieldsChanged("FindFirst:" & nField$) 'Query found
                bFindMarker = True
                bLoop = True
            End If
        End If
        
    Loop Until bLoop = True
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub FindNext()
If Connected Then
    If Len(rsLastQuery$) = 0 Then
        RaiseEvent FindNotUsed
        Exit Sub
    End If
    bFindMarker = False
    
    Do
        If dbRs(rsIndex).EOF <> True Then dbRs(rsIndex).MoveNext
        dbRs(rsIndex).Find typeCriteria$, 0, adSearchForward
            If dbRs(rsIndex).EOF Then
                If rsRequest <> 1 Then
                    If rsIndex = tblCount - 1 Then
                        bLoop = True
                    Else
                        rsIndex = rsIndex + 1
                        dbRs(rsIndex).MoveFirst
                        bLoop = False
                    End If
                Else
                    RaiseEvent QueryNotFound(rsLastQuery$)
                    bLoop = True
                End If
            Else
                If CaseSensitive = True Then
                    Dim sResult As String
                    sResult = dbRs(rsIndex).Fields(rsLastField)
                        If Trim(rsLastQuery$) = Trim(sResult) Then
                            RaiseEvent FieldsChanged("FindNext:" & rsLastField$)
                            bFindMarker = True
                            bLoop = True
                        End If
                Else
                    RaiseEvent FieldsChanged("FindNext:" & rsLastField$)
                    bFindMarker = True
                    bLoop = True
                End If
            End If
        
    Loop Until bLoop = True
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
Public Sub FindPrev()
If Connected Then
    If Len(rsLastQuery$) = 0 Then
        RaiseEvent FindNotUsed
        Exit Sub
    End If
    bFindMarker = False
    Do
        If dbRs(rsIndex).BOF <> True Then dbRs(rsIndex).MovePrevious
        dbRs(rsIndex).Find typeCriteria$, 0, adSearchBackward
            If dbRs(rsIndex).BOF Then
                If rsRequest <> 1 Then
                    If rsIndex = 0 Then
                        bLoop = True
                    Else
                        rsIndex = rsIndex - 1
                        dbRs(rsIndex).MoveLast
                    End If
                Else
                    RaiseEvent QueryNotFound(rsLastQuery$)
                    bLoop = True
                End If
            Else
                If CaseSensitive = True Then
                    Dim sResult As String
                    sResult = dbRs(rsIndex).Fields(rsLastField)
                        If Trim(rsLastQuery$) = Trim(sResult) Then
                            RaiseEvent FieldsChanged("FindPrev:" & rsLastField$)
                            bFindMarker = True
                            bLoop = True
                        End If
                Else
                    RaiseEvent FieldsChanged("FindPrev:" & rsLastField$)
                    bFindMarker = True
                    bLoop = True
                End If
            End If
        
    Loop Until bLoop = True
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If

End Sub
Public Sub FindLast(ByRef nField As String, ByRef sQuery As String, Optional ByVal RequestType As RequestType = 2)
If Connected Then
    If setCriteria(nField$, sQuery$) = False Then Exit Sub
    rsRequest = RequestType
    If RequestType <> 1 Then rsIndex = tblCount - 1
    bFindMarker = False
    
    Do
        dbRs(rsIndex).MoveLast
        dbRs(rsIndex).Find typeCriteria$, 0, adSearchBackward
                
        If dbRs(rsIndex).BOF Then
            If RequestType <> 1 Then
                If rsIndex <> 0 Then
                    rsIndex = rsIndex - 1
                Else
                    RaiseEvent QueryNotFound(sQuery$)
                    rsLastQuery$ = vbNullString
                    bLoop = True
                End If
            Else
                RaiseEvent QueryNotFound(sQuery$)
                rsLastQuery$ = vbNullString
                bLoop = True
            End If
        Else
            If CaseSensitive = True Then
                Dim sResult As String
                sResult = dbRs(rsIndex).Fields(rsLastField)
                    If Trim(sQuery) = Trim(sResult) Then
                        RaiseEvent FieldsChanged("FindLast:" & nField$)
                        bFindMarker = True
                        bLoop = True
                    Else
                        FindPrev
                        Exit Sub
                    End If
            Else
                RaiseEvent FieldsChanged("FindLast:" & nField$)
                bFindMarker = True
                bLoop = True
            End If
        End If
        
    Loop Until bLoop = True
Else
    RaiseEvent ConnectionError(errNotConnected$)
End If
End Sub
'===========================Class Functions/Subs========================================
Private Function ValidInt(qNumber As Long, ByVal iMin As Integer, ByVal iMax As Integer) As Boolean
    If qNumber Like "*[!.,0-9]*" Or qNumber Like "*.*" Then Exit Function
    If qNumber < iMin Or qNumber > iMax Then Exit Function
    ValidInt = True
End Function
Private Function RaiseErr(ByVal ErrorNumber As Long, Optional ByRef Description As String)

Select Case ErrorNumber
    Case "3265"
        RaiseEvent InvalidParameter("Field Name Unknown")
    Case Else
        RaiseEvent Error(Description$)
End Select

End Function
Private Function setCriteria(ByRef nField As String, ByRef sQuery As String) As Boolean
If Len(sQuery$) = 0 Or Len(nField$) = 0 Then
    RaiseEvent InvalidParameter("Field or Query is Invalid.")
Else
    rsLastQuery$ = sQuery$ 'update last query
    rsLastField$ = nField$ 'update last field
    typeCriteria$ = "[" & nField$ & "] = '" & sQuery$ & "'" 'set criteria for search
    setCriteria = True
    bLoop = False
End If
End Function
Private Sub checkAttrib()
On Error GoTo notWritable:
    Dim fFile As Integer
        bWritable = True
        fFile = FreeFile
    
    Open dbLocation$ For Append As #fFile
    Close #fFile
    
Exit Sub
notWritable:
    bWritable = False
End Sub
Private Function FolderExists(ByRef Location As String) As Boolean
On Error GoTo IsFalse:
    If FileSystem.GetAttr(Location) Then FolderExists = True
Exit Function
IsFalse:
End Function
Private Sub Class_Terminate()
Dispose 'Call Cleanup Code
End Sub

