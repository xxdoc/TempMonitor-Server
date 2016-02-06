VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexterSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsStatus
' Author    : XPMUser
' Date      : 2/2/2016
' Purpose   : saves a message to a file for program Texter to read. If the file has
'             not been read then the message is appended to the file. Otherwise the
'             file is replaced with the new message.
'---------------------------------------------------------------------------------------

Option Explicit
Dim Saved As Boolean
Dim Mes As String
Dim modFileName As String
Dim modTimeStamp As Boolean

Event MessageSaved()

Public Property Get FileName() As String
    FileName = modFileName
End Property

Public Property Let FileName(NewVal As String)
    modFileName = NewVal
    EraseOldData
End Property

Public Property Get TimeStamp() As Boolean
    TimeStamp = modTimeStamp
End Property

Public Property Let TimeStamp(NewVal As Boolean)
    modTimeStamp = NewVal
End Property

Private Function Save(Message As String) As Boolean
    Dim FN As Integer
    On Error GoTo ErrExit
    Save = False
    FN = FreeFile
    If (GetAttr(modFileName) And vbArchive) = 0 Then
        'file has been read, overwrite
        Open modFileName For Output Lock Read Write As #FN
    Else
        'file has not been read, append
        Open modFileName For Append Lock Read Write As #FN
    End If
    Print #FN, Message & vbNewLine
    Close #FN
    Save = True
ErrExit:
End Function

Public Sub SaveToFile(Message As String)
'---------------------------------------------------------------------------------------
' Procedure : SaveToFile
' Author    : XPMUser
' Date      : 21/Jan/2015
' Purpose   : saves message to file for Texter
' If message = "" then save old message, if any.
'---------------------------------------------------------------------------------------
'
    If TimeStamp Then
        Message = Format(Now, "dd-mmm-yy hh:mm:ss AM/PM") & ":  " & Message
    End If
    If Saved Then
        'save new message
        Mes = Message
    Else
        'add new message to old message and save
        Mes = Mes & Message
    End If
    If Mes <> "" Then
        Saved = Save(Mes)
        If Saved Then
            RaiseEvent MessageSaved
        End If
    End If
End Sub

Private Sub EraseOldData()
    Dim FN As Integer
    If Dir(modFileName) <> "" Then Kill FileName
    FN = FreeFile
    Open modFileName For Output As #FN
    Close #FN
    SetAttr modFileName, vbNormal
End Sub

Private Sub Class_Initialize()
    modTimeStamp = True
End Sub
