VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6E1B7648-872D-4B19-96AD-0555B4151387}#14.1#0"; "dcList.ocx"
Begin VB.Form frmClients 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clients"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "frmClients.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butSignal 
      Caption         =   "Check signal strength."
      Height          =   375
      Left            =   3060
      TabIndex        =   9
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox tbEvents 
      Height          =   2895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   6240
      Width           =   8055
   End
   Begin VB.CommandButton butEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      Top             =   4875
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9285
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4650
      TabIndex        =   5
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox textbox 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6510
      TabIndex        =   2
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   4875
      Width           =   855
   End
   Begin dcList.dcListControl List1 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      SortCol         =   0
      RowsToDisplay   =   8
      MinRowsToDisplay=   12
      HideSearch      =   0   'False
      GridCaption     =   "Clients"
      Fill            =   -1  'True
      MaxWidth        =   10000
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   8280
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4470
      Width           =   855
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LstObj As List_Objects
Private WithEvents objMain As clsClient
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
     AD.DisplayError Err.Number, "frmClients", "ApplyEdit", Err.Description
     Resume ErrExit
End Function

Private Sub butCancel_Click()
    On Error GoTo ErrHandler
    objMain.CancelEdit
    Editing = False
    EnableControls
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "butCancel_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butDelete_Click()
    On Error GoTo ErrHandler
    Select Case MsgBox("Confirm Delete and erase any temperature records.", vbOKCancel Or vbQuestion Or vbDefaultButton1, App.Title)
    
        Case vbOK
            Set objMain = New clsClient
            objMain.Load List1.RecordID
            If Not objMain.IsNew Then
                'record was valid, now delete it
                objMain.BeginEdit
                objMain.Delete
                objMain.ApplyEdit
            End If
            UpdateGrid
        Case vbCancel
    
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "butDelete_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butEdit_Click()
    On Error GoTo ErrHandler
    Set objMain = New clsClient
    'attempt to load current record
    objMain.Load List1.RecordID
    If objMain.IsNew Then
        'current record not valid, it stays as a new record
        EditID = -1
    Else
        EditID = objMain.ID
    End If
    objMain.BeginEdit
    Editing = True
    EnableControls
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "butEdit_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butSave_Click()
    On Error GoTo ErrHandler
    If butSave.Caption = "Close" Then
        Unload Me
    Else
        If ApplyEdit Then
            Editing = False
            EnableControls
            UpdateForm
            UpdateGrid EditID
        End If
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmClients", "butSave_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub butSignal_Click()
    On Error GoTo ErrHandler
    frmStart.SendPacket pdGetSignalStrength
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "butSignal_Click", Err.Description
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
     AD.DisplayError Err.Number, "frmClients", "CancelExit", Err.Description
     Resume ErrExit
End Function

Private Sub EnableControls()
    On Error GoTo ErrHandler
    butEdit.Enabled = Not Editing
    butCancel.Enabled = Editing
    butDelete.Enabled = Not Editing
    textbox(0).Enabled = Editing
    If Editing Then
        butSave.Caption = "Save"
    Else
        butSave.Caption = "Close"
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "EnableControls", Err.Description
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
    AD.DisplayError Err.Number, "frmClients", "EventMessage", Err.Description
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
    AD.DisplayError Err.Number, "frmClients", "Form_KeyDown", Err.Description
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
     AD.DisplayError Err.Number, "frmClients", "Form_KeyPress", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrExit
    Editing = False
    AD.LoadFormData Me
    LstObj = LOclients
    List1.RowsToDisplay = 15
    FormatGrid
    UpdateGrid
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
     AD.DisplayError Err.Number, "frmClients", "FormatGrid", Err.Description
     Resume ErrExit
End Sub

Private Sub List1_DblClick()
    butEdit_Click
End Sub

Public Sub NewClient()
    On Error GoTo ErrHandler
    If Not Editing Then UpdateGrid
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmClients", "NewSensor", Err.Description
     Resume ErrExit
End Sub

Private Sub objMain_DataEdited()
    UpdateForm
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
                .Description = textbox(Index)
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
            AD.DisplayError Err.Number, "frmClients", "textbox_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub UpdateForm()
    On Error GoTo ErrHandler
    UpdatingForm = True
    If Editing Then
        With objMain
            textbox(0) = .Description
        End With
    Else
        textbox(0) = ""
    End If
    UpdatingForm = False
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmClients", "UpdateForm", Err.Description
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
     AD.DisplayError Err.Number, "frmClients", "UpdateGrid", Err.Description
     Resume ErrExit
End Sub

