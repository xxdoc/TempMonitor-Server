VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type ObjProps
    dbType As String
    DbVersion As String
    RecordInterval As Long  'minutes
    AlarmInterval As Long   'hours (0-60) 0 means turn sensor alarm off
    TrendTime As Long       'days  (0-60) 0 means turn trend alarm off
    TrendMax As Currency    'temperature max increase
    MaxDBsize As Long       'in K
End Type

Event Valid(IsValid As Boolean)
Event DataEdited()
Event DatabaseLoaded()

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private CurrentDbName As String
Private CurrentDbPassword As String
Private mudtProps As ObjProps
Private mudtSave As ObjProps

Public Sub ApplyEdit()
    If Not mflgEditing Then Err.Raise 445
    If mflgDeleted Then
        If Not mflgNew Then
            'code to delete the object's data goes here
            'no code since object won't be deleted
            mflgNew = True
        End If
        mflgDeleted = False
    ElseIf mflgDirty Or mflgNew Then
        If Not IsValid Then Err.Raise 445
        'save object to database if appropriate
        Save
        'save object state
        LSet mudtSave = mudtProps
        mflgNew = False
    End If
    mflgDirty = False
    mflgEditing = False
End Sub

Public Sub BeginEdit()
    If mflgEditing Then Err.Raise 445
    'save object state
    LSet mudtSave = mudtProps
    mflgEditing = True
End Sub

Public Property Get BrokenRules() As BrokenRules
    Set BrokenRules = mobjValid
End Property

Public Sub CancelEdit()
    If Not mflgEditing Then Err.Raise 445
    mflgEditing = False
    mflgDeleted = False
    mflgDirty = False
    'restore object state
    LSet mudtProps = mudtSave
End Sub

Private Sub Class_Initialize()
    mflgNew = True
    Set mobjValid = New BrokenRules
    mobjValid.RuleBroken "Database", True, "A Database is required."
    mobjValid.RuleBroken "Version", True, "The database version is invalid."
    LoadDefaults
End Sub

Private Sub LoadDefaults()
    mudtProps.AlarmInterval = 12    'hours
    mudtProps.RecordInterval = 5    'minutes
    mudtProps.TrendTime = 7 'days
    mudtProps.TrendMax = 10  'degrees
    mudtProps.MaxDBsize = 750 'Kilobytes
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    MainDB.Close
End Sub

Public Property Get DataChanged() As Boolean
    DataChanged = mflgDirty
End Property

Public Property Get IsDeleted() As Boolean
    IsDeleted = mflgDeleted
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = mflgDirty Or mflgNew
End Property

Public Property Get IsNew() As Boolean
    IsNew = mflgNew
End Property

Public Property Get IsValid() As Boolean
    IsValid = (mobjValid.Count = 0)
End Property

Private Sub mobjValid_BrokenRule()
    RaiseEvent Valid(False)
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid(True)
End Sub

Private Sub Save()
    Dim SQL As String
    Dim RS As Recordset
    SQL = "select * from tblProps"
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If mflgNew Then
            .AddNew
        Else
            .Edit
        End If
        !dbType = mudtProps.dbType
        !DbVersion = mudtProps.DbVersion
        !dbRecordInterval = mudtProps.RecordInterval
        !dbAlarmInterval = mudtProps.AlarmInterval
        !dbTrendTime = mudtProps.TrendTime
        !dbTrendMax = mudtProps.TrendMax
        !dbMaxDBsize = mudtProps.MaxDBsize
        .Update
        .Close
    End With
    Set RS = Nothing
End Sub

Private Function FileExists(DBname As String) As Boolean
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    FileExists = FSO.FileExists(DBname)
    Set FSO = Nothing
End Function

Private Function CheckPassword(DBname As String, DBpassword As String) As Boolean
    Dim DB As Database
    On Error GoTo ErrExit
    CheckPassword = False
    Set DB = OpenDatabase(DBname, False, False, ";pwd=" & DBpassword)
    CheckPassword = True
ErrExit:
    Set DB = Nothing
End Function

Private Function CheckDatabaseType(DBname As String, DBpassword As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : CheckDatabase
' Author    : XPMUser
' Date      : 1/17/2016
' Purpose   : checks if database type is correct
'---------------------------------------------------------------------------------------
'
    Dim DB As Database
    Dim RS As Recordset
    On Error GoTo ErrExit
    CheckDatabaseType = False
    Set DB = OpenDatabase(DBname, False, False, ";pwd=" & DBpassword)
    Set RS = DB.OpenRecordset("Select * from tblProps")
    If LCase(RS!dbType) = LCase(AppName) Then
        CheckDatabaseType = True
    End If
ErrExit:
    Set DB = Nothing
    Set RS = Nothing
End Function

Private Function CheckDatabaseVersion(DBname As String, DBpassword As String) As Long
    Dim DB As Database
    Dim RS As Recordset
    On Error GoTo ErrExit
    CheckDatabaseVersion = 0
    Set DB = OpenDatabase(DBname, False, False, ";pwd=" & DBpassword)
    Set RS = DB.OpenRecordset("Select * from tblProps")
    CheckDatabaseVersion = RS!DbVersion
ErrExit:
    Set DB = Nothing
    Set RS = Nothing
End Function

Public Sub LoadDB(DBname As String, Optional DBpassword As String = "")
    Dim Ver As Long
    If Not FileExists(DBname) Then Err.Raise ErrFileNotFound
    If Not CheckPassword(DBname, DBpassword) Then Err.Raise ErrPassword
    If Not CheckDatabaseType(DBname, DBpassword) Then Err.Raise ErrDBtype
    Ver = CheckDatabaseVersion(DBname, DBpassword)
    If Ver < CurrentDBversion Then Err.Raise ErrLowVersion
    If Ver > CurrentDBversion Then Err.Raise ErrHighVersion
    Load DBname, DBpassword
End Sub

Private Sub Load(DBname As String, Optional DBpassword As String = "")
    Dim RS As Recordset
    Dim SQL As String
    If mflgEditing Then Err.Raise 445, "clsMain.Load"
'    If Not mflgNew Then Err.Raise 445, "clsMain.Load"
    CurrentDbName = DBname
    CurrentDbPassword = DBpassword
    'if database doesn't open, the error will raise up to
    'the calling procedure
    Set MainDB = OpenDatabase(DBname, False, False, ";pwd=" & DBpassword)
    SQL = "select * from tblProps"
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        mudtProps.dbType = LCase(NZ(!dbType, True))
        mudtProps.DbVersion = NZ(!DbVersion)
        mudtProps.RecordInterval = NZ(!dbRecordInterval)
        mudtProps.AlarmInterval = NZ(!dbAlarmInterval)
        mudtProps.TrendTime = NZ(!dbTrendTime)
        mudtProps.TrendMax = NZ(!dbTrendMax)
        mudtProps.MaxDBsize = NZ(!dbMaxDBsize)
        .Close
    End With
    mobjValid.RuleBroken "Database", (mudtProps.dbType <> LCase(AppName)), "A Database is required."
    mobjValid.RuleBroken "Version", (mudtProps.DbVersion <> CurrentDBversion), "The database version is invalid."
    mflgNew = False
    RaiseEvent DatabaseLoaded
    Set RS = Nothing
End Sub

Private Sub ValueChanged()
    mflgDirty = True
    RaiseEvent DataEdited
End Sub

'****************** fields ***********************

Public Property Get dbType() As String
    dbType = mudtProps.dbType
End Property

Public Property Get DbVersion() As String
    DbVersion = mudtProps.DbVersion
End Property

Public Property Get RecordInterval() As Long
    RecordInterval = mudtProps.RecordInterval
End Property

Public Property Let RecordInterval(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If mudtProps.RecordInterval <> NewVal Then
        If NewVal < 1 Or NewVal > 10000 Then Err.Raise InputErr, "clsMain", "Invalid Record Interval. (1 - 10,000)"
        mudtProps.RecordInterval = NewVal
        ValueChanged
    End If
End Property

Public Property Get AlarmInterval() As Long
    AlarmInterval = mudtProps.AlarmInterval
End Property

Public Property Let AlarmInterval(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If mudtProps.AlarmInterval <> NewVal Then
        If NewVal < 0 Or NewVal > 60 Then Err.Raise InputErr, "clsMain", "Invalid Alarm Interval. (0 - 60)"
        mudtProps.AlarmInterval = NewVal
        ValueChanged
    End If
End Property

Public Property Get TrendTime() As Long
    TrendTime = mudtProps.TrendTime
End Property

Public Property Let TrendTime(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If mudtProps.TrendTime <> NewVal Then
        If NewVal < 0 Or NewVal > 60 Then Err.Raise InputErr, "clsMain", "Invalid Trend Time. (0 - 60)"
        mudtProps.TrendTime = NewVal
        ValueChanged
    End If
End Property

Public Property Get TrendMax() As Currency
    TrendMax = mudtProps.TrendMax
End Property

Public Property Let TrendMax(NewVal As Currency)
    If Not mflgEditing Then Err.Raise 383
    If mudtProps.TrendMax <> NewVal Then
        If NewVal < 1 Or NewVal > 20 Then Err.Raise InputErr, "clsMain", "Invalid Trend Maximum Temperature. (1 - 20)"
        mudtProps.TrendMax = NewVal
        ValueChanged
    End If
End Property

Public Property Get MaxDBsize() As Long
    MaxDBsize = mudtProps.MaxDBsize
End Property

Public Property Let MaxDBsize(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If mudtProps.MaxDBsize <> NewVal Then
        If NewVal < 250 Or NewVal > 5000 Then Err.Raise InputErr, "clsMain", "Invalid Maximum database size. (250 - 5000)"
        mudtProps.MaxDBsize = NewVal
        ValueChanged
    End If
End Property

'****************** other properties ***********************

Public Function ChangePassword(OldPassword As String, NewPassword As String) As Boolean
    Dim DBchange As Database
    If OldPassword <> CurrentDbPassword Then Err.Raise InputErr, , "Invalid Old Password."
    'close current database
    Set MainDB = Nothing
    'have to set to exclusive with the True parameter
    'for NewPassword method to work
    Set DBchange = OpenDatabase(CurrentDbName, True, False, ";pwd=" & CurrentDbPassword)
    DBchange.NewPassword CurrentDbPassword, NewPassword
    Set DBchange = Nothing
    mflgNew = True
    Load CurrentDbName, NewPassword
    ChangePassword = True
End Function

Private Sub Shrink()
'---------------------------------------------------------------------------------------
' Procedure : Shrink
' Author    : XPMUser
' Date      : 1/16/2016
' Purpose   : remove records from tblRecords when database size is
'             greater than max
'---------------------------------------------------------------------------------------
'
    Dim FSO As New FileSystemObject
    Dim S As Long
    Dim RS As Recordset
    Dim SQL As String
    Dim ID As Long
    Dim R As Currency
    S = FSO.GetFile(CurrentDbName).Size / 1024
    If S > MaxDBsize Then
        R = S / MaxDBsize - 1
        If R > 0.5 Then R = 0.5
        'get id of record R% from start
        SQL = "select top " & Int(R * 100) & " percent * from tblrecords order by recID"
        Set RS = MainDB.OpenRecordset(SQL)
        If Not RS.EOF Then
            RS.MoveLast
            ID = RS!recID
            'delete records up to ID
            SQL = "delete * from tblrecords where RecID < " & ID
            MainDB.Execute SQL, dbFailOnError
        End If
    End If
    Set FSO = Nothing
    Set RS = Nothing
End Sub

Public Function CompactDatabase() As Boolean
    Dim Temp As String
    CompactDatabase = False
    Shrink
    'get the store folder
    Temp = MainDB.Name
    Temp = Left$(Temp, InStrRev(Temp, "\")) & "temp.mdb"
    'close any open database connections
    Workspaces(0).Close
    'erase temp name if exists
    If Dir(Temp) <> "" Then Kill Temp
    If CurrentDbPassword = "" Then
        DBEngine.CompactDatabase CurrentDbName, Temp
    Else
        DBEngine.CompactDatabase CurrentDbName, Temp, , , CurrentDbPassword
    End If
    'check if new compacted file has been created
    If Dir(Temp) <> "" Then
        'file has been created
        Kill CurrentDbName
        Name Temp As CurrentDbName
        CompactDatabase = True
    End If
    mflgNew = True
    Load CurrentDbName, CurrentDbPassword
End Function

Public Function CreateDatabase(DBname As String, TBL() As Boolean, _
    Path As String, Optional DBpassword As String) As Boolean
    
    Dim NewPath As String
    Dim RS As Recordset
    Dim DBnew As Database
    Dim RSnew As Recordset
    
    NewPath = Path
    DBname = NewPath & "\" & DBname & ".mdb"
    'check new doesn't overwrite existing
    If Dir(DBname) <> "" Or DBname = "" Then Err.Raise InputErr, , "Invalid file name."
    'create a new blank database
    If Not CreatedNewDB(DBname, DBpassword) Then Err.Raise InputErr, , "Could not create new file."
    'create reference to new database
    Set DBnew = OpenDatabase(DBname, False, False, ";pwd=" & DBpassword)
    'edit tblProps
    Set RS = DBnew.OpenRecordset("tblProps")
    With RS
        If .EOF Then
            .AddNew
        Else
            .Edit
        End If
        !dbType = AppName
        !DbVersion = CurrentDBversion
        .Update
    End With
    Set RS = Nothing
    Set RSnew = Nothing
    Set DBnew = Nothing
    mflgNew = True
    Load DBname, DBpassword
    LoadDefaults
    Save
    CreateDatabase = True
End Function

Public Function DatabaseFullName() As String
    DatabaseFullName = CurrentDbName
End Function

Public Function DatabaseName() As String
    'returns the file name out of a path
    'ex: C:\GrainManager\GrainManagerdata.mdb
    'would return 'GrainManagerdata'
    Dim P As Long
    Dim ST As Long
    Dim EN As Long
    ST = 0
    EN = Len(CurrentDbName) + 1
    For P = Len(CurrentDbName) To 1 Step -1
        If Mid$(CurrentDbName, P, 1) = "." Then EN = P
        If Mid$(CurrentDbName, P, 1) = "\" Then
            ST = P
            Exit For
        End If
    Next P
    DatabaseName = Mid$(CurrentDbName, ST + 1, EN - ST - 1)
End Function

Public Function UpdateDatabaseVersion(DBname As String, DBpassword As String)
    Dim Ver As Long
    Ver = CheckDatabaseVersion(DBname, DBpassword)
    'only update older versions, do nothing if same version
    If Ver < CurrentDBversion Then
'        'close references to database file
'        Workspaces(0).Close
        UpdateDB DBname, DBpassword, CurrentDBversion
'        mflgNew = True
'        Load CurrentDbName, CurrentDbPassword
    End If
End Function

