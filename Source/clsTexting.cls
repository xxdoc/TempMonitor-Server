VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Valid(IsValid As Boolean)
Event DataEdited()

Private mflgNew As Boolean
Private mflgDeleted As Boolean
Private mflgDirty As Boolean
Private mflgEditing As Boolean
Private WithEvents mobjValid As BrokenRules
Attribute mobjValid.VB_VarHelpID = -1

Private Type ObjProps
    ID As Long
    Port As Long
    Server As String * 30
    User As String * 30
    Password As String * 30
    FromEmail As String * 30
End Type

Private mudtProps As ObjProps
Private mudtSave As ObjProps

Public Sub ApplyEdit()
    Dim DL As String
    If Not mflgEditing Then Err.Raise 445
    If mflgDeleted Then
        If Not mflgNew Then
            'code to delete the object's data goes here
'            DL = "tblSensors,senClientID," & mudtProps.ID & ",false"
'            If DependantData(DL) Then
'                Err.Raise InputErr, "Client", "Dependant data."
'            End If
            DeleteRecord MainDB, "tblClients", mudtProps.ID, "ClientID"
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
    Dim D As String
    mflgNew = True
    Set mobjValid = New BrokenRules
    CheckRules
End Sub

Private Sub CheckRules()
    mobjValid.RuleBroken "Port", (mudtProps.Port = 0), "Invalid Port."
    mobjValid.RuleBroken "Server", (mudtProps.Server = ""), "Invalid Server."
    mobjValid.RuleBroken "User", (mudtProps.User = ""), "Invalid User name."
    mobjValid.RuleBroken "Password", (mudtProps.Password = ""), "Invalid Password."
    mobjValid.RuleBroken "FromEmail", (mudtProps.FromEmail = ""), "Invalid From Email."
End Sub

Public Property Get DataChanged() As Boolean
    DataChanged = mflgDirty
End Property

Public Sub Delete()
    If Not mflgEditing Then Err.Raise 445
    mflgDeleted = True
    mflgDirty = True
End Sub

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
    SQL = "select * from tblTexting where txtID = " & mudtProps.ID
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If mflgNew Then
            .AddNew
        Else
            .Edit
        End If
        !txtPort = mudtProps.Port
        !txtServer = mudtProps.Server
        !txtUser = mudtProps.User
        !txtPassword = mudtProps.Password
        !txtFromEmail = mudtProps.FromEmail
        .Update
        'get ID from autoincrement field after file is saved
        If mflgNew Then
            .Bookmark = .LastModified
            mudtProps.ID = !txtID
        End If
        .Close
    End With
    Set RS = Nothing
End Sub

Public Sub Load(ID As Long)
    Dim RS As Recordset
    Dim SQL As String
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    SQL = "select * from tblTexting where txtID = " & ID
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            mudtProps.ID = NZ(!txtID)
            mudtProps.Port = NZ(!txtPort)
            mudtProps.Server = NZ(!txtServer, True)
            mudtProps.User = NZ(!txtUser, True)
            mudtProps.Password = NZ(!txtPasswor, True)
            mudtProps.FromEmail = NZ(!txtFromEmail, True)
            CheckRules
            mflgNew = False
        End If
       .Close
    End With
    Set RS = Nothing
End Sub

Private Sub ValueChanged()
    mflgDirty = True
    RaiseEvent DataEdited
End Sub

'****************** fields ***********************

Friend Property Let ID(NewVal As String)
    'this is a 'Friend' property to allow the collection class
    'to work properly
    mudtProps.ID = CCur(NewVal)
End Property

Public Property Get ID() As String
    ID = mudtProps.ID
End Property

Public Property Get Port() As Long
    Port = mudtProps.Port
End Property

Public Property Let Port(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.Port Then
        If NewVal < 0 Then Err.Raise InputErr, , "Invalid port."
        mudtProps.Port = NewVal
        ValueChanged
    End If
End Property

Public Property Get Server() As String
    Server = mudtProps.Server
End Property

Public Property Let Server(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.Server Then
        If Len(NewVal) > Len(mudtProps.Server) Then Err.Raise InputErr, , "Server name too long."
        mudtProps.Server = NewVal
        ValueChanged
    End If
End Property

Public Property Get User() As String
    User = mudtProps.User
End Property

Public Property Let User(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.User Then
        If Len(NewVal) > Len(mudtProps.User) Then Err.Raise InputErr, , "User name too long."
        mudtProps.User = NewVal
        ValueChanged
    End If
End Property

Public Property Get Password() As String
    Password = mudtProps.Password
End Property

Public Property Let Password(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.Password Then
        If Len(NewVal) > Len(mudtProps.Password) Then Err.Raise InputErr, , "Password name too long."
        mudtProps.Password = NewVal
        ValueChanged
    End If
End Property

Public Property Get FromEmail() As String
    FromEmail = mudtProps.FromEmail
End Property

Public Property Let FromEmail(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.FromEmail Then
        If Len(NewVal) > Len(mudtProps.FromEmail) Then Err.Raise InputErr, , "FromEmail name too long."
        mudtProps.FromEmail = NewVal
        ValueChanged
    End If
End Property

'****************** other properties ***********************

