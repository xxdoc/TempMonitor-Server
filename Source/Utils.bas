Attribute VB_Name = "Utils"
Option Explicit
Const SND_ASYNC& = &H1
Const SND_NODEFAULT& = &H2
Const SND_NOWAIT& = &H2000
Const SND_SYNC& = &H0
Const SND_NOSTOP = &H10
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Function CheckBoxToBool(Chk As CheckBox) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : CheckBoxToBool
' Author    : David
' Date      : 1/28/2013
' Purpose   : returns true if checkbox is checked, false otherwise
'---------------------------------------------------------------------------------------
'

    On Error GoTo ErrHandler
    If Chk.Value = 0 Then
        CheckBoxToBool = False
    Else
        CheckBoxToBool = True
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Utils", "CheckBoxToBool", Err.Description
    Resume ErrExit
End Function

Private Function DatabaseExists() As Boolean
    Dim P As String
    Dim MonitorDB As Database
    On Error GoTo ErrExit
    P = AD.DataDir & "\MonitorData"
    Set MonitorDB = OpenDatabase(P)
    DatabaseExists = True
    Set MonitorDB = Nothing
ErrExit:
End Function

Private Sub DefineDB(DBname As String)
    'define the database as a Shafts database
    Dim DB As Database
    Dim RS As Recordset
    Set DB = OpenDatabase(DBname)
    Set RS = DB.OpenRecordset("tblProperties")
    With RS
        .AddNew
        !propType = "Shafts"
        !propVersion = "1.0"
        .Update
        .Close
    End With
    Set RS = Nothing
    Set DB = Nothing
End Sub

Public Function FindDB(Simulate As Boolean) As Database
'---------------------------------------------------------------------------------------
' Procedure : FindDB
' Author    : David
' Date      : 1/27/2013
' Purpose   : open database. Let errors pass up to calling procedure
'---------------------------------------------------------------------------------------
'
    Dim P As String
    If Simulate Then
        P = AD.DataDir & "\SimulateData"
        Set FindDB = OpenDatabase(P)
    Else
        If DatabaseExists Then
            P = AD.DataDir & "\MonitorData"
            Set FindDB = OpenDatabase(P)
        Else
            P = AD.DataDir & "\MonitorData"
            If CreatedNewDB(P, "") Then
                DefineDB P
                'use new database
                Set FindDB = OpenDatabase(P)
            Else
                'couldn't create database
                Err.Raise 75
            End If
        End If
    End If
End Function

Public Function GoodDB(DBname As String) As Boolean
    'checks if database is for Shaft Monitor
    Dim DB As Database
    Dim RS As Recordset
    On Error GoTo ErrExit
    Set DB = OpenDatabase(DBname)
    Set RS = DB.OpenRecordset("tblProperties")
    If RS!propType = "Shafts" Then
        GoodDB = True
    End If
ErrExit:
    On Error GoTo 0
    Set RS = Nothing
    Set DB = Nothing
End Function

Function NZ(NullValue As Variant, Optional IsString As Boolean = False) As Variant
'---------------------------------------------------------------------------------------
' Procedure : NZ
' Author    : David
' Date      : 19-Apr-09
' Purpose   : replace a null number with 0 or a null string with ""
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    NZ = NullValue
    If IsNull(NZ) Then
        If IsString Then
            NZ = ""
        Else
            NZ = 0
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Utils", "NZ", Err.Description
    Resume ErrExit
End Function

Public Sub PlaySnd(fname As String)
'---------------------------------------------------------------------------------------
' Procedure : PlaySnd
' Author    : David
' Date      : 1/4/2012
' Purpose   : .Wav sound files
'---------------------------------------------------------------------------------------
'
    Dim ret As Long
    On Error GoTo ErrHandler
    ret = sndPlaySound(fname, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP Or SND_NOWAIT)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "Utils", "PlaySnd", Err.Description
    Resume ErrExit
End Sub

Public Sub SetCheckBox(Chk As CheckBox, Value As Variant)
'---------------------------------------------------------------------------------------
' Procedure : SetCheckBox
' Date      : 15/Apr/2010
' Purpose   : convert true/false to checked/unchecked
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If Value = True Or Value = 1 Then
        Chk.Value = 1
    Else
        Chk.Value = 0
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "Utils", "SetCheckBox", Err.Description
    Resume ErrExit
End Sub

Function ToBool(V As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : ToBool
' Author    : David
' Date      : 12/24/2011
' Purpose   : converts a zero-length string to false
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If Val(V) = 0 And LCase(V) <> "true" Then
        ToBool = False
    Else
        ToBool = True
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Utils", "ToBool", Err.Description
    Resume ErrExit
End Function

