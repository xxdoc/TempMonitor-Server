VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "3"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "3"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "3"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "3"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton butSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4515
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Alarm 2: (Change in Temperature Trend)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label Label 
      Caption         =   "Alarm 1: (Maximum Sensor Temperature)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "Maximum database size  (KB)"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Trend temperature increase"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Trend time period  (days)"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Check Interval  (hours)"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Record Inverval  (minutes)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private locLoadOK As Boolean

Private Function ApplyEdit() As Boolean
    On Error GoTo ErrHandler
    Dim R As Long
    Dim Mes As String
    ApplyEdit = False
    If Prog.IsValid Then
        Prog.ApplyEdit
        ApplyEdit = True
    Else
        For R = 1 To Prog.BrokenRules.Count
            Mes = Mes & Prog.BrokenRules.RuleDescription(R) & " "
        Next R
        StatusBar1.SimpleText = Mes
        Beep
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmInvoices", "ApplyEdit", Err.Description
     Resume ErrExit
End Function

Private Sub butCancel_Click()
    On Error GoTo ErrHandler
    Prog.CancelEdit
    Prog.BeginEdit
    UpdateForm
    butSave.Caption = "Close"
    butCancel.Enabled = False
    textbox(0).SetFocus
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmOptions", "butCancel_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub butSave_Click()
    On Error GoTo ErrHandler
    If butSave.Caption = "Close" Then
        Unload Me
    Else
        If ApplyEdit Then
            Unload Me
        End If
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmOptions", "butSave_Click", Err.Description
    Resume ErrExit
End Sub

Private Function CancelExit() As Boolean
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
     AD.DisplayError Err.Number, "frmOptions", "CancelExit", Err.Description
     Resume ErrExit
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    Select Case KeyCode
        Case 38
            'up arrow
            SendKeys ("+{tab}")
        Case 40
            'down arrow
            SendKeys ("{tab}")
    End Select
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmOptions", "Form_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'enter
        SendKeys ("{tab}")
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrExit
    AD.LoadFormData Me
    butSave.Caption = "Close"
    butCancel.Enabled = False
    StatusBar1.SimpleText = ""
    Prog.BeginEdit
    UpdateForm
    locLoadOK = True
ErrExit:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ValidateControls
    If Err = 380 Then
        'control validation event set to cancel
        Cancel = True
    Else
        If Prog.DataChanged Then Cancel = CancelExit
    End If
    If Not Cancel Then
        AD.SaveFormData Me
        Prog.CancelEdit
    End If
End Sub

Public Property Get LoadOk() As Boolean
    LoadOk = locLoadOK
End Property

Private Sub textbox_GotFocus(Index As Integer)
    textbox(Index).SelStart = 0
    textbox(Index).SelLength = Len(textbox(Index).Text)
End Sub

Private Sub textbox_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo ErrHandler
    Dim ER As Long
    StatusBar1.SimpleText = ""
    With Prog
        Select Case Index
            Case 0
                .RecordInterval = CLng(textbox(Index))
            Case 1
                .MaxDBsize = CLng(textbox(Index))
            Case 2
                .AlarmInterval = CLng(textbox(Index))
            Case 3
                .TrendTime = CLng(textbox(Index))
            Case 4
                .TrendMax = CCur(textbox(Index))
        End Select
    End With
    If Prog.DataChanged Then
        UpdateForm
        butSave.Caption = "Save"
        butCancel.Enabled = True
    End If
ErrExit:
    Exit Sub
ErrHandler:
    'convert err number to GrainManager object error
    ER = (Err.Number And &HFFFF&)
    Select Case ER
        Case 1001
            'GrainManager object input error
            Beep
            StatusBar1.SimpleText = Err.Description
            Cancel = True
            UpdateForm
            textbox_GotFocus Index
        Case Else
            AD.DisplayError Err.Number, "frmOptions", "textbox_Validate", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub UpdateForm()
    On Error GoTo ErrHandler
    With Prog
        textbox(0).Text = .RecordInterval
        textbox(1).Text = .MaxDBsize
        textbox(2).Text = .AlarmInterval
        textbox(3).Text = .TrendTime
        textbox(4).Text = .TrendMax
    End With
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmOptions", "UpdateForm", Err.Description
    Resume ErrExit
End Sub

