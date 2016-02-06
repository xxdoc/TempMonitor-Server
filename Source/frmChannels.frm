VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChannels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Channels"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   Icon            =   "frmChannels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdSheet 
      Height          =   7600
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13414
      _Version        =   393216
      Rows            =   10
      Cols            =   8
      FixedCols       =   2
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Press <Delete> to delete as sensor."
      Top             =   8010
      Width           =   14310
      _ExtentX        =   25241
      _ExtentY        =   873
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChannels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Data fields
'chID
'chChannel
'chDesciption
'chType
'chEnabled
'chDisplayID
'chAlarmEnabled
'chMinimum
'chMagnets

Const MaxRows As Integer = 20

Private WithEvents mSheet As clsSheet
Attribute mSheet.VB_VarHelpID = -1
Private Invalid As Boolean
Private ID(MaxRows) As Long
Private LastCol As Long
Private modSimulate As Boolean

Private Sub DeleteRow()
    Dim DB As Database
    Dim SQL As String
    Dim RS As Recordset
    Dim Rw As Integer
    On Error GoTo ErrHandler
    Set DB = FindDB(modSimulate)
    Rw = mSheet.Row
    SQL = "select * from tblChannels where chID = " & ID(Rw)
    Set RS = DB.OpenRecordset(SQL)
    With RS
        If Not .EOF Then .Delete
    End With
    Set RS = Nothing
    Set DB = Nothing
    grdSheet.Clear
    SetUpSheet
    LoadData
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "DeleteRow", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Load()
    LastCol = -1
    SetUpSheet
    LoadData
    grdSheet_RowColChange
    AD.LoadFormData Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'check current row
    mSheet.Validate
    Cancel = Invalid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub

Private Sub grdSheet_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    Select Case KeyCode
        Case 46
            'delete
            Select Case MsgBox("Confirm delete sensor?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.Title)
                Case vbYes
                    DeleteRow
                    Form_Load
            End Select
            KeyCode = 0
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "grdSheet_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub grdSheet_RowColChange()
    Dim Mes As String
    On Error GoTo ErrHandler
    With grdSheet
        If .Col <> LastCol Then
            LastCol = .Col
            Select Case .Col - 1
                Case 1
                    'description
                    Mes = "Maximum of 12 characters"
                Case 2
                    'type
                    Mes = "0 (Reed Switch), 1 (Relay)"
                Case 3
                    'enabled
                    Mes = "1 (True) or 0 (False)"
                Case 4
                    'DisplayID
                    Mes = "0 to 19"
                Case 5
                    'AlarmEnabled
                    Mes = "1 (True) or 0 (False)"
                Case 6
                    'Minimum
                    Mes = "0 to 100,000"
                Case 7
                    'magnets
                    Mes = "1 to 10"
                Case 8
                    'relay delay
                    Mes = "0 to 900 seconds"
            End Select
            StatusBar1.SimpleText = Mes
        End If
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "grdSheet_RowColChange", Err.Description
    Resume ErrExit
End Sub

Private Sub LoadData()
    Dim DB As Database
    Dim RS As Recordset
    Dim R As Integer
    On Error GoTo ErrHandler
    Set DB = FindDB(modSimulate)
    Set RS = DB.OpenRecordset("tblChannels")
    With RS
        Do Until .EOF
            R = NZ(!chChannel) + 1
            If R <= MaxRows Then
                mSheet.Cell(R, 1) = NZ(!chDescription, True)
                mSheet.Cell(R, 2) = NZ(!chType)
                mSheet.Cell(R, 3) = NZ(!chEnabled)
                mSheet.Cell(R, 4) = NZ(!chDisplayID)
                mSheet.Cell(R, 5) = NZ(!chAlarmEnabled)
                mSheet.Cell(R, 6) = NZ(!chMinimum)
                mSheet.Cell(R, 7) = NZ(!chMagnets)
                mSheet.Cell(R, 8) = NZ(!chDelay)
                ID(R) = NZ(!chID)
            End If
            .MoveNext
        Loop
    End With
    Set RS = Nothing
    Set DB = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "LoadData", Err.Description
    Resume ErrExit
End Sub

Private Sub mSheet_ValidateRow(OldRow As Integer, ReturnCol As Integer, Cancel As Boolean)
    Dim C As Integer
    Dim V As Currency
    Dim S As String
    Dim IsRelay As Boolean
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    With mSheet
        IsRelay = (Val(.Cell(OldRow, 2)) = 1)
        For C = 1 To 8
            V = Val(.Cell(OldRow, C))
            S = .Cell(OldRow, C)
            Select Case C
                Case 1
                    'description
                    If .Cell(OldRow, 1) = "" Then
                        StatusBar1.SimpleText = "Invalid Description."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                    If Len(.Cell(OldRow, 1)) > 12 Then
                        StatusBar1.SimpleText = "Description too long. Maximum is 12 characters."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 2
                    'type
                    If (V <> 0 And V <> 1) Or Not IsNumeric(S) Then
                        StatusBar1.SimpleText = "Invalid Type value. Should be either 0 Reed Switch or 1 Relay."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 3
                    'enabled
                    If Not IsRelay And ((V <> 0 And V <> 1) Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Invalid Enabled value. Should be either 1 True or 0 False."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 4
                    'displayID
                    If Not IsRelay And (V < 0 Or V > 19 Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Invalid DisplayID #. Should be 0 to 19."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 5
                    'AlarmEnabled
                    If Not IsRelay And ((V <> 0 And V <> 1) Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Invalid AlarmEnabled value. Should be either 1 True or 0 False."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 6
                    'minimum
                    If Not IsRelay And (V < 0 Or V > 100000 Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Minimum RPM. Should be 0 to 100000."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 7
                    'Magnets
                    If Not IsRelay And (V < 1 Or V > 10 Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Invalid # of Magnets. Should be 1 to 10."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
                Case 8
                    'delay
                    If IsRelay And (V < 0 Or V > 900 Or Not IsNumeric(S)) Then
                        StatusBar1.SimpleText = "Invalid Relay Delay. Should be 0 to 900 seconds."
                        Cancel = True
                        ReturnCol = C
                        Exit For
                    End If
            End Select
        Next C
        If Not Cancel Then SaveData OldRow
    End With
    Invalid = Cancel
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "mSheet_ValidateRow", Err.Description
    Resume ErrExit
End Sub

Private Sub SaveData(Rw As Integer)
    Dim DB As Database
    Dim SQL As String
    Dim RS As Recordset
    Dim IsNew As Boolean
    On Error GoTo ErrHandler
    Set DB = FindDB(modSimulate)
    SQL = "select * from tblChannels where chID = " & ID(Rw)
    Set RS = DB.OpenRecordset(SQL)
    With RS
        If .EOF Then
            .AddNew
            IsNew = True
        Else
            .Edit
        End If
        !chDescription = mSheet.Cell(Rw, 1)
        !chType = Val(mSheet.Cell(Rw, 2))
        !chEnabled = Val(mSheet.Cell(Rw, 3))
        !chDisplayID = Val(mSheet.Cell(Rw, 4))
        !chAlarmEnabled = Val(mSheet.Cell(Rw, 5))
        !chMinimum = Val(mSheet.Cell(Rw, 6))
        !chMagnets = Val(mSheet.Cell(Rw, 7))
        !chDelay = Val(mSheet.Cell(Rw, 8))
        !chChannel = Rw - 1
        .Update
        If IsNew Then
            .Bookmark = .LastModified
            ID(Rw) = NZ(!chID)
        End If
    End With
    Set RS = Nothing
    Set DB = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "SaveData", Err.Description
    Resume ErrExit
End Sub

Private Sub SetUpSheet()
    Dim R As Integer
    Dim W As Long
    Dim H As Long
    On Error GoTo ErrHandler
    Set mSheet = New clsSheet           ' Create the new class
    With mSheet
        Set .Grid = grdSheet            ' Assign MSFlexGrid and TextBox controls
        Set .EditBox = txtEdit
        .Rows = MaxRows                      ' Set the sheet size
        .Cols = 8
        For R = 1 To MaxRows
            .RowTitle(R) = R - 1
        Next R
        .ColTitle(0) = "Channel"
        .ColWidth(0) = 1000
        .ColAlign(0) = flexAlignCenterCenter
        .ColTitle(1) = "Description"
        .ColWidth(1) = 2000
        .ColAlign(1) = flexAlignLeftCenter
        .ColTitle(2) = "Type"
        .ColAlign(2) = flexAlignCenterCenter
        .ColTitle(3) = "Enabled"
        .ColAlign(3) = flexAlignCenterCenter
        .ColTitle(4) = "DisplayID"
        .ColAlign(4) = flexAlignCenterCenter
        .ColWidth(4) = 1400
        .ColTitle(5) = "AlarmEnabled"
        .ColAlign(5) = flexAlignCenterCenter
        .ColWidth(5) = 1800
        .ColTitle(6) = "Minimum RPM"
        .ColAlign(6) = flexAlignCenterCenter
        .ColWidth(6) = 1800
        .ColTitle(7) = "Magnets"
        .ColAlign(7) = flexAlignCenterCenter
        .ColTitle(8) = "Delay"
        .ColAlign(8) = flexAlignCenterCenter
    End With
    grdSheet.Width = mSheet.GridWidth + 40
    H = StatusBar1.Height + grdSheet.Height + 460
    W = grdSheet.Width + 100
    If H > (Screen.Height - 100) Then H = Screen.Height - 100
    If W > (Screen.Width - 100) Then W = Screen.Width - 100
    Me.Height = H
    Me.Width = W
    If grdSheet.Width > (W - 100) Then grdSheet.Width = W - 100
    If grdSheet.Height > (H - 100) Then grdSheet.Height = H - 100
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmChannels", "SetUpSheet", Err.Description
    Resume ErrExit
End Sub

Public Property Let Simulate(NewVal As Boolean)
    modSimulate = NewVal
End Property

