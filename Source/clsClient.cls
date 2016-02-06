VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    RecNum As Long
    Mac As String * 30
    Description As String * 20
    SocketID As Integer
End Type

Private mudtProps As ObjProps
Private mudtSave As ObjProps

Public Sub ApplyEdit()
    Dim DL As String
    If Not mflgEditing Then Err.Raise 445
    If mflgDeleted Then
        If Not mflgNew Then
            'code to delete the object's data goes here
            DL = "tblSensors,senClientID," & mudtProps.ID & ",false"
            If DependantData(DL) Then
                Err.Raise InputErr, "Client", "Dependant data."
            End If
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
    mflgNew = True
    Set mobjValid = New BrokenRules
    
    'Enter default values for new records
    mudtProps.RecNum = CCur(LastValue(MainDB, "tblClients", "ClientRecNum", "ClientRecNum")) + 1
    
    'initialize the list of broken rules
    mobjValid.RuleBroken "Mac", TR(mudtProps.Mac) = "", "Invalid Mac address."
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
    SQL = "select * from tblClients where ClientID = " & mudtProps.ID
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If mflgNew Then
            .AddNew
        Else
            .Edit
        End If
        !ClientRecNum = mudtProps.RecNum
        !ClientMac = TR(mudtProps.Mac)
        !ClientDescription = TR(mudtProps.Description)
        !ClientSocketID = mudtProps.SocketID
        .Update
        'get ID from autoincrement field after file is saved
        If mflgNew Then
            .Bookmark = .LastModified
            mudtProps.ID = !ClientID
        End If
        .Close
    End With
    Set RS = Nothing
End Sub

Public Sub Load(Optional ID As Long, Optional Mac As String, Optional Sckt As Integer)
    Dim RS As Recordset
    Dim SQL As String
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    If Mac <> "" Then
        'load by Mac address
        SQL = "select * from tblClients where ClientMac = '" & Mac & "'"
    ElseIf Sckt > 0 Then
        'load by socket
        SQL = "select * from tblClients where ClientSocketID = " & Sckt
    Else
        'load by ID
        SQL = "select * from tblClients where ClientID = " & ID
    End If
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            mudtProps.ID = NZ(!ClientID)
            mudtProps.RecNum = NZ(!ClientRecNum)
            mudtProps.Mac = NZ(!ClientMac, True)
            mudtProps.Description = NZ(!ClientDescription, True)
            mudtProps.SocketID = NZ(!ClientSocketID)
            mobjValid.RuleBroken "Mac", False
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

Friend Property Let RecNum(NewVal As String)
    mudtProps.RecNum = CCur(NewVal)
End Property

Public Property Get RecNum() As String
    RecNum = mudtProps.RecNum
End Property

Public Property Get Description() As String
    Description = TR(mudtProps.Description)
    If Description = "" Then Description = TR(mudtProps.Mac)
End Property

Public Property Let Description(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If TR(mudtProps.Description) <> TR(NewVal) Then
        If Len(NewVal) > Len(mudtProps.Description) Then
            Err.Raise InputErr, "", "Description too long."
        End If
        mudtProps.Description = NewVal
        ValueChanged
    End If
End Property

Public Property Get Mac() As String
    Mac = TR(mudtProps.Mac)
End Property

Public Property Let Mac(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If TR(mudtProps.Mac) <> TR(NewVal) Then
        If Len(NewVal) > Len(mudtProps.Mac) Then
            Err.Raise InputErr, "clsClient", "Mac address too long."
        End If
        mudtProps.Mac = NewVal
        ValueChanged
        mobjValid.RuleBroken "Mac", TR(mudtProps.Mac) = "", "Invalid Mac address."
    End If
End Property

Public Property Get SocketID() As Integer
    SocketID = mudtProps.SocketID
End Property

Public Property Let SocketID(NewVal As Integer)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.SocketID Then
        mudtProps.SocketID = NewVal
        ValueChanged
    End If
End Property

'****************** other properties ***********************
