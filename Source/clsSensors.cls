VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSensors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private modCol As Collection
Private Sub Class_Initialize()
    Set modCol = New Collection
End Sub
Public Function Count() As Long
    Count = modCol.Count
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'set procedure attributes ID = -4, Hidden checked
    Set NewEnum = modCol.[_NewEnum]
End Function
Public Function Item(ByVal Index As Variant) As clsSensor
Attribute Item.VB_UserMemId = 0
    'set procedure attributes ID = default
    'find or build an object
    'then returnit as a result
    'change the 'Object' to the name of the object
    'ex: Item(ByVal Index As Variant) As Customer
    Set Item = modCol(Index)
End Function
Public Sub Load(Optional ID As Long, Optional Bin As Long, Optional SensorNumber As Long)
    'enter optional parameters for this sub
    'ex: Load(Optional Name as string, Optional Phone as string)
    'load data from database
    Dim RS As Recordset
    Dim SQL As String
    Dim modObject As clsSensor
    If ID > 0 Then
        SQL = SQL & " where senID = " & ID
    End If
    If Bin > 0 Then
        If SQL = "" Then
            SQL = SQL & " where "
        Else
            SQL = SQL & " and "
        End If
        SQL = SQL & "senStorID = " & Bin
    End If
    If SensorNumber > 0 Then
        If SQL = "" Then
            SQL = SQL & " where "
        Else
            SQL = SQL & " and "
        End If
        SQL = SQL & "senNumber = " & SensorNumber
    End If
    SQL = "Select * from tblSensors" & SQL
    Set RS = MainDB.OpenRecordset(SQL)
    Do While Not RS.EOF
        Set modObject = New clsSensor
        modObject.Load RS!SenID
        modCol.Add modObject, Format$(modObject.ID) & "ID"
        Set modObject = Nothing
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    Set modObject = Nothing
End Sub



