VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSensor"
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
    Description As String * 20
    RomCode As String * 30
    senNumber As Long
    BinID As Long
    Status As SensorStatus
    senType As SensorTypes
    ClientID As Long
    MaxTemp As Currency
    TextInterval As Long 'days
    LastText As Date    'date of last text message successfully sent
    DailyTemp As Currency
    TrendTemp As Currency
End Type

Private mudtProps As ObjProps
Private mudtSave As ObjProps

Public Sub ApplyEdit()
    Dim DL As String
    If Not mflgEditing Then Err.Raise 445
    If mflgDeleted Then
        If Not mflgNew Then
            'code to delete the object's data goes here
            DL = "tblRecords,recSenID," & mudtProps.ID & ",false"
'            DL = DL & ",tblSettings,SetCropID," & mudtProps.ID & ",false"
            If DependantData(DL) Then
                Err.Raise InputErr, "Sensor", "Dependant data."
            End If
            DeleteRecord MainDB, "tblSensors", mudtProps.ID, "senID"
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
    mudtProps.RecNum = CCur(LastValue(MainDB, "tblSensors", "senRecNum")) + 1
    mudtProps.senNumber = 1
    mudtProps.Status = ssEnabled
    mudtProps.senType = stTemperature
    mudtProps.MaxTemp = 40
    'initialize the list of broken rules
    CheckRules
End Sub

Private Sub CheckRules()
    mobjValid.RuleBroken "RomCode", Not ValidMacAddress(TR(mudtProps.RomCode)), "Invalid Rom Code."
'    mobjValid.RuleBroken "Number", Not ValidSensorNumber(mudtProps.senNumber, mudtProps.BinID), "Invalid sensor number."
'    mobjValid.RuleBroken "BinNumber", (mudtProps.BinID = 0), "Invalid bin number."
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
    SQL = "select * from tblSensors where senID = " & mudtProps.ID
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If mflgNew Then
            .AddNew
        Else
            .Edit
        End If
        !senRecNum = mudtProps.RecNum
        !senDescription = TR(mudtProps.Description)
        !senMac = mudtProps.RomCode
        !senNumber = mudtProps.senNumber
        !senStorID = mudtProps.BinID
        !senStatus = mudtProps.Status
        !senType = mudtProps.senType
        !senClientID = mudtProps.ClientID
        !senMaxTemp = mudtProps.MaxTemp
        !senTextInterval = mudtProps.TextInterval
        !senLastText = mudtProps.LastText
        !senDailyTemp = mudtProps.DailyTemp
        !senTrendTemp = mudtProps.TrendTemp
        .Update
        'get ID from autoincrement field after file is saved
        If mflgNew Then
            .Bookmark = .LastModified
            mudtProps.ID = !SenID
        End If
        .Close
    End With
    Set RS = Nothing
End Sub

Public Sub Load(SenID As Long, Optional RomCode As String)
    Dim RS As Recordset
    Dim SQL As String
    If mflgEditing Then Err.Raise 445
    If Not mflgNew Then Err.Raise 445
    If RomCode = "" Then
        SQL = "select * from tblSensors where senID = " & SenID
    Else
        SQL = "select * from tblSensors where senMac = '" & RomCode & "'"
    End If
    Set RS = MainDB.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            mudtProps.Description = NZ(!senDescription, True)
            mudtProps.RomCode = NZ(!senMac)
            mudtProps.senNumber = NZ(!senNumber)
            mudtProps.BinID = NZ(!senStorID)
            mudtProps.Status = NZ(!senStatus)
            mudtProps.senType = NZ(!senType)
            mudtProps.ID = NZ(!SenID)
            mudtProps.RecNum = NZ(!senRecNum)
            mudtProps.ClientID = NZ(!senClientID)
            mudtProps.MaxTemp = NZ(!senMaxTemp)
            mudtProps.TextInterval = NZ(!senTextInterval)
            mudtProps.LastText = NZ(!senLastText)
            mudtProps.DailyTemp = NZ(!senDailyTemp)
            mudtProps.TrendTemp = NZ(!senTrendTemp)
            mflgNew = False
            CheckRules
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

Public Property Get RomCode() As String
    RomCode = TR(mudtProps.RomCode)
End Property

Public Property Let RomCode(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If TR(mudtProps.RomCode) <> TR(NewVal) Then
        If Len(NewVal) > Len(mudtProps.RomCode) Then
            Err.Raise InputErr, "", "Rom Code too long."
        End If
        mudtProps.RomCode = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get Number() As String
    Number = mudtProps.senNumber
End Property

Public Property Let Number(NewVal As String)
    If Not mflgEditing Then Err.Raise 383
    If Not IsNumeric(NewVal) Then
        Err.Raise InputErr, "clsSensor", "A number is required."
    End If
    If mudtProps.senNumber <> CLng(NewVal) Then
        mudtProps.senNumber = CLng(NewVal)
        CheckRules
        ValueChanged
    End If

End Property

Public Property Get BinID() As String
    BinID = mudtProps.BinID
End Property

Public Property Let BinID(NewVal As String)
    Dim NV As Long
    If Not mflgEditing Then Err.Raise 383
    If Not IsNumeric(NewVal) Then
        Err.Raise InputErr, "BinTest", "A number is required."
    End If
    NV = CLng(NewVal)
    If NV <> mudtProps.BinID Then
        'check NewVal exists
        If Not ValueFound(MainDB, "tblStorage", "StorID", NV) Then
            Err.Raise InputErr, "clsSensor", "Bin not found."
        End If
        mudtProps.BinID = NV
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get Status() As SensorStatus
    Status = mudtProps.Status
End Property

Public Property Let Status(NewVal As SensorStatus)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.Status Then
        mudtProps.Status = NewVal
        ValueChanged
    End If
End Property

Public Property Get senType() As SensorTypes
    senType = mudtProps.senType
End Property

Public Property Let senType(NewVal As SensorTypes)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.senType Then
        mudtProps.senType = NewVal
        ValueChanged
    End If
End Property

Public Property Get ClientID() As Long
    ClientID = mudtProps.ClientID
End Property

Public Property Let ClientID(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.ClientID Then
        'check if exists, but allow setting it to 0
        If Not ValueFound(MainDB, "tblClients", "ClientID", NewVal) And NewVal <> 0 Then
            Err.Raise InputErr, "clsSensor", "Client not found."
        End If
        mudtProps.ClientID = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get MaxTemp() As Currency
    MaxTemp = mudtProps.MaxTemp
End Property

Public Property Let MaxTemp(NewVal As Currency)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.MaxTemp Then
        If NewVal < 0 Or NewVal > 50 Then Err.Raise InputErr, "clsSensor", "Invalid maximum temperature."
        mudtProps.MaxTemp = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get TextInterval() As Long
    TextInterval = mudtProps.TextInterval
End Property

Public Property Let TextInterval(NewVal As Long)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.TextInterval Then
        If NewVal < 0 Or NewVal > 60 Then Err.Raise InputErr, "clsSensor", "Invalid Text Interval. (0 - 60)"
        mudtProps.TextInterval = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get LastText() As Date
    LastText = mudtProps.LastText
End Property

Public Property Let LastText(NewVal As Date)
    If Not mflgEditing Then Err.Raise 383
    If Int(NewVal) <> Int(mudtProps.LastText) Then
        If Int(NewVal) > Int(Now) Then Err.Raise InputErr, "clsSensor", "Invalid date."
        mudtProps.LastText = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get DailyTemp() As Currency
    DailyTemp = Format(mudtProps.DailyTemp, "###.0")
End Property

Public Property Let DailyTemp(NewVal As Currency)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.DailyTemp Then
        If NewVal < -50 Or NewVal > 50 Then Err.Raise InputErr, "clsSensor", "Invalid temperature."
        mudtProps.DailyTemp = NewVal
        CheckRules
        ValueChanged
    End If
End Property

Public Property Get TrendTemp() As Currency
    TrendTemp = Format(mudtProps.TrendTemp, "###.0")
End Property

Public Property Let TrendTemp(NewVal As Currency)
    If Not mflgEditing Then Err.Raise 383
    If NewVal <> mudtProps.TrendTemp Then
        If NewVal < -50 Or NewVal > 50 Then Err.Raise InputErr, "clsSensor", "Invalid temperature."
        mudtProps.TrendTemp = NewVal
        CheckRules
        ValueChanged
    End If
End Property

'****************** other properties ***********************

Public Function ValidSensorNumber(NewNumber As Long, BinID As Long) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : ValidSensorNumber
' Author    : XPMUser
' Date      : 12/16/2015
' Purpose   : make sure bin number is only used once per bin
'---------------------------------------------------------------------------------------
'
    Dim SQL As String
    Dim RS As Recordset
    Dim Count As Long
    SQL = "select * from tblSensors where senStorID = " & BinID
    SQL = SQL & " and senNumber = " & NewNumber
    Set RS = MainDB.OpenRecordset(SQL)
    Do Until RS.EOF
        Count = Count + 1
        RS.MoveNext
    Loop
    ValidSensorNumber = (Count < 2)
    Set RS = Nothing
End Function

Public Function BinNumber() As String
    Dim Bin As Storage
    If ValueFound(MainDB, "tblStorage", "StorID", mudtProps.BinID) Then
        Set Bin = New Storage
        Bin.Load (mudtProps.BinID)
        BinNumber = Bin.Label
        Set Bin = Nothing
    End If
End Function

Public Function ValidMacAddress(NewMac As String) As Boolean
    Dim SQL As String
    Dim RS As Recordset
    Dim Count As Long
    If NewMac = "" Then
        ValidMacAddress = False
    Else
        SQL = "select * from tblSensors where senMac = '" & NewMac & "'"
        Set RS = MainDB.OpenRecordset(SQL)
        Do Until RS.EOF
            Count = Count + 1
            RS.MoveNext
        Loop
        If mflgNew Then
            'new object, should be none in db
            ValidMacAddress = (Count = 0)
        Else
            'editing an object, should be only 1 in db
            ValidMacAddress = (Count = 1)
        End If
    End If
    Set RS = Nothing
End Function

Public Function BinDescription() As String
    Dim Obj As Storage
    If mudtProps.BinID > 0 Then
        Set Obj = New Storage
        Obj.Load mudtProps.BinID
        BinDescription = Obj.Label & "  " & Obj.Description
    End If
    Set Obj = Nothing
End Function

Public Sub ClearRecords()
    Dim RS As Recordset
    Dim SQL As String
    SQL = "select * from tblRecords where recSenID = " & mudtProps.ID
    Set RS = MainDB.OpenRecordset(SQL)
    Do Until RS.EOF
        RS.Delete
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Public Function ClientDescription() As String
    Dim RS As Recordset
    Dim SQL As String
    SQL = "select * from tblClients where ClientID = " & mudtProps.ClientID
    Set RS = MainDB.OpenRecordset(SQL)
    If Not RS.EOF Then
        ClientDescription = TR(NZ(RS!ClientDescription))
        If ClientDescription = "" Then
            ClientDescription = TR(NZ(RS!ClientMac))
        End If
    End If
    Set RS = Nothing
End Function

Public Function TempLastDay() As Currency
'---------------------------------------------------------------------------------------
' Procedure : TempLastDay
' Author    : XPMUser
' Date      : 1/23/2016
' Purpose   : returns average temp of last 1 day for sensor
'---------------------------------------------------------------------------------------
'
    Dim SQL As String
    Dim RS As Recordset
    Dim LastDate As Date
    SQL = "select * from tblRecords where recSenID = " & mudtProps.ID
    SQL = SQL & " order by recDate"
    Set RS = MainDB.OpenRecordset(SQL)
    If Not RS.EOF Then
        RS.MoveLast
        LastDate = Int(RS!recDate)
        SQL = "select avg(recTemp) as Temp from tblRecords"
        SQL = SQL & " where int(recDate) = " & ToAccessDate(LastDate)
        SQL = SQL & " and recSenID = " & mudtProps.ID
        Set RS = MainDB.OpenRecordset(SQL)
        TempLastDay = Format(NZ(RS!Temp), "###.0")
    End If
    Set RS = Nothing
End Function

Public Function Trend(Interval As Long) As Currency
'---------------------------------------------------------------------------------------
' Procedure : Trend
' Author    : XPMUser
' Date      : 1/23/2016
' Purpose   : returns average temp over trend time period for sensor
'---------------------------------------------------------------------------------------
'
    Dim SQL As String
    Dim RS As Recordset
    Dim LastDate As Date
    SQL = "select * from tblRecords where recSenID = " & mudtProps.ID
    SQL = SQL & " order by recDate"
    Set RS = MainDB.OpenRecordset(SQL)
    If Not RS.EOF Then
        RS.MoveLast
        LastDate = DateAdd("d", -Interval, Int(RS!recDate))
        SQL = "select avg(recTemp) as Temp from tblRecords"
        SQL = SQL & " where int(recDate) = " & ToAccessDate(LastDate)
        SQL = SQL & " and recSenID = " & mudtProps.ID
        Set RS = MainDB.OpenRecordset(SQL)
        If IsNull(RS!Temp) Then
            'no data
            Trend = 0
        Else
            Trend = Format(DailyTemp - NZ(RS!Temp), "###.0")
        End If
    End If
    Set RS = Nothing

End Function
