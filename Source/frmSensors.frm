VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D9794188-4A19-4061-A8C1-BCC9E392E6DB}#3.1#0"; "dcCombo.ocx"
Object = "{6E1B7648-872D-4B19-96AD-0555B4151387}#14.1#0"; "dcList.ocx"
Begin VB.Form frmSensors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sensors"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
   Icon            =   "frmSensors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butMaxTemp 
      Caption         =   "Set maximum temperature for all sensors"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   5475
      Width           =   3135
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   10680
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "0 disables"
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton butNumbers 
      Caption         =   "Query sensor numbers."
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton butQuery 
      Caption         =   "Get Temperatures"
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox tbEvents 
      Height          =   3255
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   6840
      Width           =   11535
   End
   Begin VB.CommandButton butEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   5475
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   10335
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin dcCombo.dcComboControl Combo 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   4350
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7860
      TabIndex        =   10
      Top             =   5475
      Width           =   855
   End
   Begin VB.TextBox textbox 
      Height          =   285
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   4395
      Width           =   2535
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9900
      TabIndex        =   5
      Top             =   5475
      Width           =   855
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   5475
      Width           =   855
   End
   Begin VB.TextBox textbox 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   4860
      Width           =   975
   End
   Begin dcList.dcListControl List1 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
      SortCol         =   0
      RowsToDisplay   =   8
      MinRowsToDisplay=   12
      HideSearch      =   0   'False
      GridCaption     =   "Sensors"
      Fill            =   -1  'True
      MaxWidth        =   15000
   End
   Begin VB.Label Label6 
      Caption         =   "Maximum Temperature Alarm"
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   4875
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   11760
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   4410
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Sensor #"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4875
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Bin Number"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   4410
      Width           =   975
   End
End
Attribute VB_Name = "frmSensors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LstObj As List_Objects
Private WithEvents objMain As clsSensor
Attribute objMain.VB_VarHelpID = -1
Private UpdatingForm As Boolean
Private Editing As Boolean
Private EditID As Long

Private Function ApplyEdit() As Boolean
    On Error GoTo ErrHandler
    Dim R As Long
    Dim Mes As String
    ApplyEdit = False
    If objMain.IsValid Then
        objMain.ApplyEdit
        ApplyEdit = True
    Else
        For R = 1 To objMain.BrokenRules.Count
            Mes = Mes & objMain.BrokenRules.RuleDescription(R) & " "
        Next R
        StatusBar1.SimpleText = Mes
        Beep
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "ApplyEdit", Err.Description
     Resume ErrExit
End Function

Private Sub butCancel_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    objMain.CancelEdit
    Editing = False
    EnableControls
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "butCancel_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butDelete_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    Select Case MsgBox("Confirm Delete and erase any temperature records.", vbOKCancel Or vbQuestion Or vbDefaultButton1, App.Title)
    
        Case vbOK
            'erase any temp records
            objMain.ClearRecords
            objMain.BeginEdit
            objMain.Delete
            objMain.ApplyEdit
            UpdateForm
            UpdateGrid
        Case vbCancel
    
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "butDelete_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butEdit_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    objMain.BeginEdit
    Editing = True
    EnableControls
    Combo(0).SetFocus
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "butEdit_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butMaxTemp_Click()
    Dim Col As clsSensors
    Dim Obj As clsSensor
    Dim Ans As String
    Dim ER As Long
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    Set Col = New clsSensors
    Col.Load
    Ans = InputBox("Maximum temperature for all sensors? (Enter 0 to disable.)")
    If Ans <> "" Then
        For Each Obj In Col
            With Obj
                .BeginEdit
                .MaxTemp = CLng(Ans)
                .ApplyEdit
            End With
        Next
        UpdateGrid
    End If
    Set Col = Nothing
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    'convert err number to object error
    ER = (Err.Number And &HFFFF&)
    Select Case ER
        Case 1001
            'object input error
            Beep
            StatusBar1.SimpleText = Err.Description
        Case Else
            AD.DisplayError Err.Number, "frmSensors", "butMaxTemp_Click", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub butNumbers_Click()
    On Error GoTo ErrExit
    StatusBar1.SimpleText = ""
    frmStart.SendPacket pdGetBinNum, , List1.RecordID
    TimesUp
    While Not TimesUp(2)
        DoEvents
    Wend
    frmStart.SendPacket pdGetSenNum, , List1.RecordID
ErrExit:
End Sub

Private Sub butQuery_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    frmStart.SendPacket pdGetTemperatures
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "butQuery_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butSave_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    If butSave.Caption = "Close" Then
        Unload Me
    Else
        If ApplyEdit Then
            Editing = False
            EnableControls
            UpdateForm
            UpdateGrid EditID
            SaveNumbers
        End If
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSensors", "butSave_Click", Err.Description
    Resume ErrExit
End Sub

Private Function CancelExit() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : CancelExit
' Author    : XPMUser
' Date      : 24/Jan/2015
' Purpose   : checks if users wants to save changes before exiting the form
'---------------------------------------------------------------------------------------
'
    On Error GoTo ErrHandler
    Select Case MsgBox("Do you want to save the changes?", vbYesNoCancel Or _
        vbExclamation Or vbDefaultButton1, App.Title)
        Case vbYes
            'attempt to save before exit
            CancelExit = Not ApplyEdit
        Case vbNo
            'exit without saving
            CancelExit = False
        Case vbCancel
            'cancel exit
            CancelExit = True
    End Select
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "CancelExit", Err.Description
     Resume ErrExit
End Function

Private Sub Combo_Add(Index As Integer)
'---------------------------------------------------------------------------------------
' Procedure : Combo_Add
' Author    : XPMUser
' Date      : 1/24/2016
' Purpose   : show object's form so new item can be added
'---------------------------------------------------------------------------------------
'
    Dim FN As String
    On Error GoTo ErrHandler
    FN = FormName(Combo(Index).ObjectID)
    ShowForm FN
    ComboLoad
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "Combo_Add", Err.Description
     Resume ErrExit
End Sub

Private Sub Combo_Change(Index As Integer)
    If Not UpdatingForm Then Combo_Validate Index, False
End Sub

Private Sub Combo_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo ErrHandler
    Dim ER As Long
    StatusBar1.SimpleText = ""
    'check for no selection
    If Combo(Index).ListIndex <> -1 Then
        With objMain
            Select Case Index
                Case 0
                    .BinID = Combo(Index).RecordID
            End Select
        End With
    End If
ErrExit:
    Exit Sub
ErrHandler:
    'convert err number to FarmManager object error
    ER = (Err.Number And &HFFFF&)
    Select Case ER
        Case 1001
            'object input error
            Beep
            StatusBar1.SimpleText = Err.Description
            Cancel = True
            UpdateForm
        Case Else
            AD.DisplayError Err.Number, "frmSensors", "Combo_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub ComboLoad()
    On Error GoTo ErrHandler
    Dim Obj As StorageDisplay
    Dim Col As Storages
    'bins
    Combo(0).Clear
    Set Col = New Storages
    Col.Load , , , GMSTBins
    For Each Obj In Col
        Combo(0).AddItem Obj.Label & "  " & Obj.Description, Obj.ID
    Next
    Combo(0).ObjectID = List_Objects.LObins
    Set Obj = Nothing
    Set Col = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSensors", "ComboLoad", Err.Description
    Resume ErrExit
End Sub

Private Sub EnableControls()
    On Error GoTo ErrHandler
    butEdit.Enabled = Not Editing
    butCancel.Enabled = Editing
    butDelete.Enabled = Not Editing
    Combo(0).Enabled = Editing
    textbox(0).Enabled = Editing
    textbox(1).Enabled = Editing
    textbox(2).Enabled = Editing
    butMaxTemp.Enabled = Not Editing
    If Editing Then
        butSave.Caption = "Save"
    Else
        butSave.Caption = "Close"
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "EnableControls", Err.Description
     Resume ErrExit
End Sub

Public Sub EventMessage(ByVal Mes As String, Optional Procedure As String, _
    Optional ShowTime As Boolean = True)
'---------------------------------------------------------------------------------------
' Procedure : EventMessage
' Author    : David
' Date      : 3/12/2012
' Purpose   : send status message to the user
'---------------------------------------------------------------------------------------
'
    Dim L As Long
    On Error GoTo ErrHandler
    L = Len(tbEvents)
    If L > 5000 Then
        tbEvents.Text = Right$(tbEvents.Text, 1000)
    End If
    If ShowTime Then
        Mes = Format(Now, "hh:mm:ss  AM/PM") & "    " & Mes
    End If
    tbEvents.Text = tbEvents.Text & Mes & vbNewLine
    tbEvents.SelStart = Len(tbEvents.Text)
    AD.SaveToLog Mes
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSensors", "EventMessage", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    Dim N As Long
    N = -1
    If Me.ActiveControl.Name = "textbox" Then
        N = Me.ActiveControl.Index
    End If
    Select Case N
        Case 8
            'do nothing to allow movement in the
            'multiline textbox
        Case Else
            Select Case KeyCode
                Case 38
                    'up arrow
                    SendKeys ("+{tab}")
                Case 40
                    'down arrow
                    SendKeys ("{tab}")
            End Select
    End Select
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSensors", "Form_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim N As Long
    On Error GoTo ErrHandler
    N = -1
    If Me.ActiveControl.Name = "textbox" Then
        N = Me.ActiveControl.Index
    End If
    Select Case N
        Case 8
            'do nothing
        Case Else
            If KeyAscii = 13 Then
                'enter
                SendKeys ("{tab}")
                KeyAscii = 0
            End If
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "Form_KeyPress", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrExit
    Set objMain = New clsSensor
    Editing = False
    ComboLoad
    AD.LoadFormData Me
    LstObj = LOsensors
    List1.RowsToDisplay = 15
    FormatGrid
    UpdateGrid
    List1_RecordIdChange
    StatusBar1.SimpleText = ""
    UpdateForm
    EnableControls
ErrExit:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ValidateControls
    If Err = 380 Then
        'control validation event set to cancel
        Cancel = True
    Else
        If Editing Then
            If objMain.DataChanged Then Cancel = CancelExit
        End If
    End If
    If Not Cancel Then
        AD.SaveFormData Me
        Set objMain = Nothing
    End If
End Sub

Private Sub FormatGrid()
    Dim Temp() As String
    On Error GoTo ErrHandler
    With List1
        .Caption = ListGridCaption(LstObj)
        Temp = ListData(LstObj, 1)
        .ColProps = Temp(0)
        Temp = ListData(LstObj, 2)
        .Searches = Temp(0)
        Temp = ListData(LstObj, 4)
        .SortCol = Val(Temp(0))
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "FormatGrid", Err.Description
     Resume ErrExit
End Sub

Private Sub List1_DblClick()
    butEdit_Click
End Sub

Private Sub List1_RecordIdChange()
    On Error GoTo ErrHandler
    Set objMain = New clsSensor
    objMain.Load List1.RecordID
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "List1_RecordIdChange", Err.Description
     Resume ErrExit
End Sub

Private Sub objMain_DataEdited()
    UpdateForm
End Sub

Private Sub SaveNumbers()
    Dim Obj As clsSensor
    On Error GoTo ErrExit
    Set Obj = New clsSensor
    Obj.Load List1.RecordID
    frmStart.SendPacket pdSetBinNum, Obj.BinNumber, List1.RecordID
    TimesUp
    While Not TimesUp(2)
        DoEvents
    Wend
    frmStart.SendPacket pdSetSenNum, Obj.Number, List1.RecordID
ErrExit:
    Set Obj = Nothing
End Sub

Public Sub SensorsChanged()
    On Error GoTo ErrHandler
    If Not Editing Then UpdateGrid
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "SensorsChanged", Err.Description
     Resume ErrExit
End Sub

Private Sub textbox_GotFocus(Index As Integer)
    textbox(Index).SelStart = 0
    textbox(Index).SelLength = Len(textbox(Index).Text)
End Sub

Private Sub textbox_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo ErrHandler
    Dim ER As Long
    StatusBar1.SimpleText = ""
    With objMain
        Select Case Index
            Case 0
                .Number = textbox(Index)
            Case 1
                .Description = textbox(Index)
            Case 2
                .MaxTemp = CLng(textbox(Index))
            Case 3
                .TextInterval = CLng(textbox(Index))
        End Select
    End With
ErrExit:
    Exit Sub
ErrHandler:
    'convert err number to object error
    ER = (Err.Number And &HFFFF&)
    Select Case ER
        Case 1001
            'object input error
            Beep
            StatusBar1.SimpleText = Err.Description
            Cancel = True
            UpdateForm
            textbox_GotFocus Index
        Case Else
            AD.DisplayError Err.Number, "frmSensors", "textbox_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub UpdateForm()
    On Error GoTo ErrHandler
    UpdatingForm = True
    With objMain
        Combo(0).MoveToRecord .BinID
        textbox(0) = .Number
        textbox(1) = .Description
        textbox(2) = .MaxTemp
    End With
    UpdatingForm = False
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSensors", "UpdateForm", Err.Description
    Resume ErrExit
End Sub

Private Sub UpdateGrid(Optional MoveToRecordID As Long = -1)
    On Error GoTo ErrHandler
    With List1
        .GridData = ListData(LstObj, 3)
        .LoadGridData
        .MoveToRecord dcToRecord, MoveToRecordID
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSensors", "UpdateGrid", Err.Description
     Resume ErrExit
End Sub

