VERSION 5.00
Begin VB.Form frmClassDemo 
   Caption         =   "ezDatabase Class Application Demo"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect To Database"
      Height          =   375
      Left            =   7560
      TabIndex        =   49
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Frame fraInfo 
      Caption         =   "ezDatabase Information Commands"
      Height          =   2175
      Left            =   120
      TabIndex        =   31
      Top             =   5040
      Width           =   9615
      Begin VB.TextBox txtTableNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8760
         TabIndex        =   48
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdTableName 
         Caption         =   "Table Name By Number"
         Height          =   375
         Left            =   6720
         TabIndex        =   47
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdTable 
         Caption         =   "Table Count"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdCase 
         Caption         =   "Case Sensitive"
         Height          =   375
         Left            =   4920
         TabIndex        =   39
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdRequest 
         Caption         =   "Last Request Type"
         Height          =   375
         Left            =   3240
         TabIndex        =   45
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdPassword 
         Caption         =   "Get Password"
         Height          =   375
         Left            =   4920
         TabIndex        =   44
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdUsername 
         Caption         =   "Get Username"
         Height          =   375
         Left            =   3240
         TabIndex        =   43
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdCTNumber 
         Caption         =   "Current Table Num"
         Height          =   375
         Left            =   4920
         TabIndex        =   42
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdCTName 
         Caption         =   "Current Table Name"
         Height          =   375
         Left            =   3240
         TabIndex        =   41
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdConnected 
         Caption         =   "Connection State"
         Height          =   375
         Left            =   4920
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdDatabase 
         Caption         =   "Database Location"
         Height          =   375
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRecordCount 
         Caption         =   "Record Count"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdEOF 
         Caption         =   "End Of File"
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cboType3 
         Height          =   360
         ItemData        =   "frmClassDemo.frx":0000
         Left            =   120
         List            =   "frmClassDemo.frx":000A
         TabIndex        =   34
         Text            =   "Select Information Type (Optional)"
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton cmdBOF 
         Caption         =   "Beginning Of File"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblInfo 
         Height          =   975
         Left            =   7200
         TabIndex        =   57
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblCL 
         Caption         =   "Current Location:"
         Height          =   255
         Left            =   7200
         TabIndex        =   56
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame fraMove 
      Caption         =   "ezDatabase Move/Update/AddNew/Delete Functions"
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   9615
      Begin VB.CommandButton cmdLastTable 
         Caption         =   "Last Table"
         Height          =   375
         Left            =   1440
         TabIndex        =   53
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrevTable 
         Caption         =   "Previous Table"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdNextTable 
         Caption         =   "Next Table"
         Height          =   375
         Left            =   1440
         TabIndex        =   51
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdFirstTable 
         Caption         =   "First Table"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Record"
         Height          =   375
         Left            =   8160
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   1
         Left            =   6480
         TabIndex        =   29
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Index           =   0
         Left            =   3720
         TabIndex        =   27
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6480
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDomain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3720
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   7560
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboType2 
         Height          =   360
         ItemData        =   "frmClassDemo.frx":0029
         Left            =   120
         List            =   "frmClassDemo.frx":0033
         TabIndex        =   19
         Text            =   "Select Move Type (Optional)"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdMoveLast 
         Caption         =   "Move Last"
         Height          =   375
         Left            =   6480
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdMovePrev 
         Caption         =   "Move Previous"
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdMoveNext 
         Caption         =   "Move Next"
         Height          =   375
         Left            =   3960
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdMoveFirst 
         Caption         =   "Move First"
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblTableName 
         Caption         =   "Current Table:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblIPAdd 
         Alignment       =   2  'Center
         Caption         =   "Domain's IP:"
         Height          =   255
         Left            =   5400
         TabIndex        =   28
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label lblDomainAdd 
         Caption         =   "Domain Name:"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label lblDomainIP 
         Alignment       =   2  'Center
         Caption         =   "Domain's IP:"
         Height          =   255
         Left            =   5400
         TabIndex        =   24
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label lblDomain 
         Caption         =   "Domain Name:"
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   280
         Width           =   1095
      End
   End
   Begin VB.ComboBox cboType 
      Height          =   360
      ItemData        =   "frmClassDemo.frx":0062
      Left            =   240
      List            =   "frmClassDemo.frx":006C
      TabIndex        =   9
      Text            =   "---- Select Find Type (Optional) ----"
      Top             =   840
      Width           =   2895
   End
   Begin VB.Frame fraFind 
      Caption         =   "ezDatabase Find Functions"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last Query"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdResult 
         Caption         =   "Query Result"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto-clear Results on Query"
         Height          =   615
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Results"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ListBox lstResults 
         Height          =   1500
         Left            =   3120
         TabIndex        =   11
         Top             =   1200
         Width           =   6375
      End
      Begin VB.CommandButton cmdFindAll 
         Caption         =   "Find All"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdFindLast 
         Caption         =   "Find Last"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdFindPrev 
         Caption         =   "Find Previous"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdFindFirst 
         Caption         =   "Find First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkCase 
         Appearance      =   0  'Flat
         Caption         =   "Case Sensitive"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7920
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtQuery 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Text            =   "yahoo.com"
         Top             =   240
         Width           =   3495
      End
      Begin VB.ComboBox cboField 
         Height          =   360
         ItemData        =   "frmClassDemo.frx":009B
         Left            =   120
         List            =   "frmClassDemo.frx":00A5
         TabIndex        =   1
         Text            =   "-------- Select Field to Query -------"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblQuery 
         Caption         =   "Find Record:"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   280
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmClassDemo.frx":00BB
      Height          =   1095
      Left            =   120
      TabIndex        =   55
      Top             =   7320
      Width           =   7335
   End
End
Attribute VB_Name = "frmClassDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Information :::::::::::::::::::::::::::::::::::::::::::::::'
'Before using this application, you must reference one of   '
'Microsoft's ActiveX Data Object Library DLL's. I used v2.8 '
'but I'm sure most of the older versions will work as well. '
'To reference it, click "Project" > "References" and find it'
'there. Enjoy the free source code, this demo is fully comm-'
'ented. Also please vote or leave comments on the PSC page. '
'::::::::::::::::::::::::::::::::::: Thanks for Downloading '

Private WithEvents dns As ezDatabase 'Bind ezDatabase events to dns
Attribute dns.VB_VarHelpID = -1
Private bType As Byte

Private Sub cboField_Click()
'Enable the buttons.
cmdFindFirst.Enabled = True
cmdFindNext.Enabled = True
cmdFindPrev.Enabled = True
cmdFindLast.Enabled = True
cmdFindAll.Enabled = True
End Sub

Private Sub cboField_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
If txtField(0).Text Or txtField(1).Text = "" Then
    MsgBox "Fields must contain data."
    If txtField(0).Text = "" Then txtField(0).SetFocus Else txtField(1).SetFocus
Else
    
    Dim bCategory As Byte 'Categorize Domain Name so we know which table to add record
        bCategory = (Asc(LCase(txtField(0))) - 97) \ 4
        If bCategory > 4 Then bCategory = 4
        
        'tblNumber = Table Number to Add Record
        'objArray = Object Array to retrieve field information from
        'iStart = Starting index of array to retrieve field information from (Optional)
        dns.AddNew bCategory, txtField 'Add the record
End If
End Sub

Private Sub cmdBOF_Click()
'Property Access: Read Only
SetOptions 3 'Set Options (Request Type)
MsgBox dns.BOF(bType) 'Is database BOF?
End Sub

Private Sub cmdCase_Click()
'Property Access: Read/Write
MsgBox dns.CaseSensitive 'Returns True/False Boolean
End Sub

Private Sub cmdClear_Click()
lstResults.Clear
End Sub

Private Sub cmdConnect_Click()
If cmdConnect.Caption = "Connect To Database" Then
    'Set Database Location
    dns.dbLocation = App.Path & "\" & "DNSdb.mdb"
    
    'Add tables (recordsets)
    dns.AddTable "AD"   'Each AddTable call creates a Recordset used to
    dns.AddTable "EH"   'query information from the appropriate table.
    dns.AddTable "IL"
    dns.AddTable "MP"
    dns.AddTable "QZ"
    
    'Make Connections
    dns.Expose
    cmdConnect.Caption = "Disconnect From Database"
Else
    'Close Connections
    dns.Dispose
    cmdConnect.Caption = "Connect To Database"
End If
End Sub

Private Sub cmdConnected_Click()
'Property Access: Read Only
MsgBox dns.Connected 'Returns True/False Boolean
End Sub

Private Sub cmdCTName_Click()
'Property Access: Read Only
MsgBox dns.CurrentTable 'Returns String Value for Current Table Name
End Sub

Private Sub cmdCTNumber_Click()
'Property Access: Read Only
MsgBox dns.TableNumber 'Returns Long Value for Current Table Number
End Sub

Private Sub cmdDatabase_Click()
'Property Access: Read/Write
MsgBox dns.dbLocation 'Returns String for location of .mdb file
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete the current record?", _
            vbYesNo, "Delete Record") = vbYes Then
    dns.Delete 'Delete current record
Else
    'Do nothing
End If

End Sub

Private Sub cmdEOF_Click()
'Property Access: Read Only
SetOptions 3 'Set Options (Request Type)
MsgBox dns.EOF(bType) 'Is database BOF?
End Sub

Private Sub cmdFindAll_Click()
SetOptions 1 'Set Options (Case Sensitive/Request Type)
dns.FindAll cboField.Text, txtQuery.Text, bType
End Sub

Private Sub cmdFindFirst_Click()
SetOptions 1 'Set Options (Case Sensitive/Request Type)
'nField = Field to search
'sQuery = Query to find
'RequestType = Search either current table or all tables
dns.Find cboField.Text, txtQuery.Text, bType 'Find is equivalent to FindFirst
End Sub

Private Sub cmdFindLast_Click()
SetOptions 1 'Set Options (Case Sensitive/Request Type)
dns.FindLast cboField.Text, txtQuery.Text, bType
End Sub

Private Sub cmdFindNext_Click()
SetOptions 1 'Set Options (Case Sensitive/Request Type)
dns.FindNext 'Find Next - Find First/Last must be used before Find Next will work.
UpdateInfo
End Sub

Private Sub cmdFindPrev_Click()
SetOptions 1 'Set Options (Case Sensitive/Request Type)
dns.FindPrev 'Find Previous - Find First/Last must be used before Find Next will work.
UpdateInfo
End Sub

Private Sub cmdFirstTable_Click()
dns.FirstRecordset 'Move to first table/recordset
End Sub

Private Sub cmdLast_Click()
MsgBox dns.LastQuery
End Sub

Private Sub cmdLastTable_Click()
dns.LastRecordset 'Move to Last Recordset/Table
End Sub

Private Sub cmdMoveFirst_Click()
SetOptions 2 'Set options for Move
dns.MoveFirst bType 'MoveFirst in Current or All Tables? All by default
End Sub

Private Sub cmdMoveLast_Click()
SetOptions 2 'Set options for Move
dns.MoveLast bType 'MoveLast in Current or All Tables? All by default
End Sub

Private Sub cmdMoveNext_Click()
SetOptions 2 'Set options for Move
dns.MoveNext 'Can only work if MoveFirst/MoveLast is first used.
UpdateInfo
End Sub

Private Sub cmdMovePrev_Click()
dns.MovePrev 'Can only work if MoveFirst/MoveLast is first used.
UpdateInfo
End Sub

Private Sub cmdNextTable_Click()
dns.NextRecordset 'Moves to next table/recordset
End Sub

Private Sub cmdPassword_Click()
'Property Access: Read/Write
MsgBox dns.Password 'Returns String Value for Password
End Sub

Private Sub cmdPrevTable_Click()
dns.PreviousRecordset 'Move to Previous Recordset/Table
End Sub

Private Sub cmdRecordCount_Click()
'Property Access: Read Only
SetOptions 3 'Set Options (Request Type)

If bType = 1 Then
    MsgBox dns.RecordCount(rtCurrentTable) 'RecordCount of Current Table
Else
    MsgBox dns.RecordCount(rtAllTables) 'RecordCount of All Tables
End If
End Sub

Private Sub cmdRequest_Click()
'Property Access: Read Only
MsgBox dns.RequestType 'Returns Long Value for Last Request Type
End Sub

Private Sub cmdResult_Click()
'FindMarker property determines whether the find function was successful in finding
'a result. Returns a true/false boolean.

MsgBox dns.FindMarker
End Sub

Private Sub cmdTable_Click()
'Property Access: Read Only
MsgBox dns.TableCount 'Returns Long Value expressing recordset/table count
End Sub

Private Sub cmdTableName_Click()
'Property Access: Read Only
MsgBox dns.TableName(txtTableNum.Text)  'Returns String Value for Table Name
                                        'by Table Number
End Sub

Private Sub cmdUpdate_Click()
'rsField = Field to Update
'rsNewValue = Updated Value for Field
dns.Update "Domain", txtDomain.Text
dns.Update "DomainIP", txtIP.Text
End Sub

Private Sub cmdUsername_Click()
'Property Access: Read/Write
MsgBox dns.Username 'Returns String Value for Username
End Sub


Private Sub dns_Added()
'Event Triggered When: AddNew function is successfully used.
MsgBox "Record Added!"
End Sub

Private Sub dns_Connect()
'Event Triggered When: Connected to Database
MsgBox "Successfully Connected to Database"
End Sub

Private Sub dns_ConnectionError(Description As String)
'Event Triggered When: Could not connect to Database.
MsgBox Description
End Sub

Private Sub dns_Deleted()
'Event Triggered When: Record is successfully deleted.
MsgBox "Record Deleted!"
End Sub

Private Sub dns_Disconnect()
'Event Triggered When: Disconnected from Database.
MsgBox "Thanks for testing out my ezDatabase Class!"
End Sub

Private Sub dns_Error(Description As String)
'Event Triggered When: An unspecific/uncommon error has occured.
MsgBox Description
End Sub

Private Sub dns_FieldsChanged(QueryInfo As String)
'Event Triggered When: Find/Move/Add/Delete/Update functions are used.
If Left(QueryInfo$, 4) = "Find" Then            'If Fields Changed by Find Function,
    If cboField.Text = "Domain" Then            'Add to List
        lstResults.AddItem dns.Field("DomainIP")
    Else
        lstResults.AddItem dns.Field("Domain")
    End If
ElseIf Left(QueryInfo$, 4) = "Move" Then        'If Fields Changed by Move Function,
    txtDomain.Text = dns.Field("Domain")        'Update TextBoxes
    txtIP.Text = dns.Field("DomainIP")
    lblTableName.Caption = "Current Table: " & dns.CurrentTable 'Update Table Name
End If
UpdateInfo
End Sub

Private Sub dns_FindNotUsed()
'Event Triggered When: Find Next/Find Previous is used but FindFirst/FindLast have
'not yet been used.
MsgBox "Find First/Last Not yet used!"
End Sub

Private Sub dns_InvalidParameter(Description As String)
'Event Triggered When: An Invalid Parameter was passed to any of the functions.
MsgBox Description
End Sub

Private Sub dns_QueryNotFound(Query As String)
'Event Triggered When: Find Function was used but no results are returned.
MsgBox Query & " not found."
End Sub

Private Sub dns_Updated()
'Event Triggered When: A record is updated.
MsgBox "Record Updated!"
End Sub

Private Sub Form_Load()
Set dns = New ezDatabase
End Sub
Private Function UpdateInfo()
    lblInfo.Caption = dns.CurrentTable & " (" & dns.TableNumber & ") - " & _
                        "Item: " & dns.Field(0) & vbNewLine & _
                        "BOF: " & dns.BOF(rtCurrentTable) & vbNewLine & "EOF: " & _
                        dns.EOF(rtCurrentTable)
End Function
Private Function SetOptions(qType As Byte)
If qType = 1 Then 'Find
    If chkCase.Value = 1 Then       'Set CaseSensitive to True.
        dns.CaseSensitive = True    'Query to find must be an EXACT match.
    Else
        dns.CaseSensitive = False   'Set CaseSensitive to False
    End If
    
    If Left(cboType.Text, 2) = "--" Then
        bType = 2 'Set type of search to default (Search all tables)
    Else
        If InStr(cboType.Text, "Current") Then
            bType = 1 'Search Current Table
        Else
            bType = 2 'Search All Tables
        End If
    End If
    
    If chkAuto.Value = 1 Then lstResults.Clear 'Clear results for next query
ElseIf qType = 2 Then 'Move
    If Left(cboType2.Text, 6) = "Select" Then
        bType = 2 'Set type of search to default (Search all tables)
    Else
        If InStr(cboType2.Text, "Current") Then
            bType = 1 'Search Current Table
        Else
            bType = 2 'Search All Tables
        End If
    End If
Else
    If Left(cboType3.Text, 6) = "Select" Then
        bType = 2 'Set type of search to default (Search all tables)
    Else
        If InStr(cboType3.Text, "Current") Then
            bType = 1 'Search Current Table
        Else
            bType = 2 'Search All Tables
        End If
    End If
End If
End Function

