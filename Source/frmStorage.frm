VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D9794188-4A19-4061-A8C1-BCC9E392E6DB}#3.1#0"; "dcCombo.ocx"
Object = "{6E1B7648-872D-4B19-96AD-0555B4151387}#14.1#0"; "dcList.ocx"
Begin VB.Form frmStorage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Storage Locations"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   Icon            =   "frmStorage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton butNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4770
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton butEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   3780
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton butDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2790
      TabIndex        =   7
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "3"
      Top             =   4170
      Width           =   3135
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6225
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "3"
      Top             =   4650
      Width           =   3135
   End
   Begin dcCombo.dcComboControl Combo 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
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
   Begin dcList.dcListControl List1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   7223
      SortCol         =   0
      RowsToDisplay   =   8
      MinRowsToDisplay=   12
      HideSearch      =   0   'False
      GridCaption     =   "Maps"
      Fill            =   -1  'True
      MaxWidth        =   15000
   End
   Begin VB.Label Label4 
      Caption         =   "Storage Map"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Location #"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "frmStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LstObj As List_Objects
Private WithEvents objMain As Storage
Attribute objMain.VB_VarHelpID = -1
Private UpdatingForm As Boolean
Private Editing As Boolean
Private EditID As Long
Private modLoadOK As Boolean

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
     AD.DisplayError Err.Number, "frmStorage", "ApplyEdit", Err.Description
     Resume ErrExit
End Function

Private Sub butCancel_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    objMain.CancelEdit
    Editing = False
    EnableControls
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStorage", "butCancel_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butDelete_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    Select Case MsgBox("Confirm Delete?", vbOKCancel Or vbQuestion Or vbDefaultButton1, App.Title)
    
        Case vbOK
            objMain.BeginEdit
            objMain.Delete
            objMain.ApplyEdit
            EditID = 0
            UpdateForm
            UpdateGrid
        Case vbCancel
    
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStorage", "butDelete_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butEdit_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    objMain.BeginEdit
    Editing = True
    EnableControls
    textbox(0).SetFocus
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStorage", "butEdit_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butNew_Click()
    On Error GoTo ErrHandler
    StatusBar1.SimpleText = ""
    Set objMain = New Storage
    UpdateForm
    objMain.BeginEdit
    Editing = True
    EnableControls
    textbox(0).SetFocus
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStorage", "butNew_Click", Err.Description
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
            UpdateGrid objMain.ID
        End If
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStorage", "butSave_Click", Err.Description
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
     AD.DisplayError Err.Number, "frmStorage", "CancelExit", Err.Description
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
     AD.DisplayError Err.Number, "frmStorage", "Combo_Add", Err.Description
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
                    .MapID = Combo(Index).RecordID
            End Select
        End With
    End If
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
        Case Else
            AD.DisplayError Err.Number, "frmStorage", "Combo_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub ComboLoad()
    On Error GoTo ErrHandler
    Dim Obj As MapDisplay
    Dim Col As Maps
    'Maps
    Combo(0).Clear
    Set Col = New Maps
    Col.Load
    For Each Obj In Col
        Combo(0).AddItem Obj.MapName, Obj.ID
    Next
    Combo(0).ObjectID = List_Objects.LOmaps
    Set Obj = Nothing
    Set Col = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmBinTests", "ComboLoad", Err.Description
    Resume ErrExit
End Sub

Private Sub EnableControls()
    On Error GoTo ErrHandler
    butNew.Enabled = Not Editing
    butEdit.Enabled = Not Editing
    butCancel.Enabled = Editing
    butDelete.Enabled = Not Editing
    textbox(0).Enabled = Editing
    textbox(1).Enabled = Editing
    Combo(0).Enabled = Editing
    If Editing Then
        butSave.Caption = "Save"
    Else
        butSave.Caption = "Close"
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStorage", "EnableControls", Err.Description
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
    AD.DisplayError Err.Number, "frmStorage", "Form_KeyDown", Err.Description
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
     AD.DisplayError Err.Number, "frmStorage", "Form_KeyPress", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrExit
    Set objMain = New Storage
    ComboLoad
    Editing = False
    AD.LoadFormData Me
    LstObj = LObins
    List1.RowsToDisplay = 15
    FormatGrid
    UpdateGrid
    StatusBar1.SimpleText = ""
    UpdateForm
    EnableControls
    modLoadOK = True
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
     AD.DisplayError Err.Number, "frmStorage", "FormatGrid", Err.Description
     Resume ErrExit
End Sub

Private Sub List1_DblClick()
    butEdit_Click
End Sub

Public Property Get LoadOK() As Boolean
    LoadOK = modLoadOK
End Property

Private Sub List1_RecordIdChange()
    UpdateForm
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
                .Label = textbox(Index)
            Case 1
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
            AD.DisplayError Err.Number, "frmStorage", "textbox_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub UpdateForm()
    On Error GoTo ErrHandler
    UpdatingForm = True
    If List1.RecordID <> EditID And Editing = False Then
        Set objMain = New Storage
        objMain.Load List1.RecordID
        EditID = List1.RecordID
    End If
    With objMain
        textbox(0) = .Label
        textbox(1) = .Description
        Combo(0).MoveToRecord .MapID
    End With
    UpdatingForm = False
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStorage", "UpdateForm", Err.Description
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
     AD.DisplayError Err.Number, "frmStorage", "UpdateGrid", Err.Description
     Resume ErrExit
End Sub

