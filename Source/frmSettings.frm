VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7605
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   2160
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton butDefaults 
      Caption         =   "Load Defaults"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   7335
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1320
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   2010
      Width           =   1320
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   1455
      Width           =   1320
   End
   Begin VB.TextBox textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   915
      Width           =   1320
   End
   Begin VB.Label Label10 
      Caption         =   "for clients to connect to (1 to 65000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   15
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   "Server Port"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "% of average RPM (0-100)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   11
      Top             =   2085
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "# of scans per state change (0-100)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   10
      Top             =   1530
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "millisecond interval to read stream (1-5000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   9
      Top             =   990
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "# of scans/second (1-5000)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   8
      Top             =   435
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Scan Rate"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Alarm Set"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2010
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Debounce"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1456
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Read Interval"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   915
      Width           =   1695
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FldCount As Long = 5
Private CurrentTextBox As Integer
Private ScanRate As Long
Private ReadInterval As Long
Private Debounce As Long
Private AlarmSet As Long
Private Port As Long

Private Sub butDefaults_Click()
    On Error GoTo ErrHandler
    ScanRate = 1500
    ReadInterval = 1000
    Debounce = 2
    AlarmSet = 60
    Port = 1234
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "butDefaults_Click", Err.Description
    Resume ErrExit
End Sub
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
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "Form_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        'enter
        SendKeys ("{tab}")
        KeyAscii = 0
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "Form_KeyPress", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    AD.LoadFormData Me
    LoadData
    UpdateForm
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "Form_Load", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Textbox_Validate CurrentTextBox, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    AD.SaveFormData Me
    SaveData
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "Form_Unload", Err.Description
    Resume ErrExit
End Sub

Private Sub LoadData()
    On Error GoTo ErrHandler
    'check if this is a new file
    If AD.AppData("Defaults") = "" Then
        butDefaults_Click
        SaveData
        AD.AppData("Defaults") = "True"
    Else
        ScanRate = Val(AD.AppData("ScanRate"))
        ReadInterval = Val(AD.AppData("ReadInterval"))
        Debounce = Val(AD.AppData("DeBounce"))
        AlarmSet = Val(AD.AppData("AlarmSet"))
        Port = Val(AD.AppData("LocalPort"))
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "LoadData", Err.Description
    Resume ErrExit
End Sub

Private Sub SaveData()
    On Error GoTo ErrHandler
    AD.AppData("ScanRate") = ScanRate
    AD.AppData("ReadInterval") = ReadInterval
    AD.AppData("DeBounce") = Debounce
    AD.AppData("AlarmSet") = AlarmSet
    AD.AppData("LocalPort") = Port
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "SaveData", Err.Description
    Resume ErrExit
End Sub

Private Sub textbox_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    textbox(Index).SelStart = 0
    textbox(Index).SelLength = Len(textbox(Index).Text)
    CurrentTextBox = Index
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "textbox_GotFocus", Err.Description
    Resume ErrExit
End Sub

Private Sub Textbox_Validate(Index As Integer, Cancel As Boolean)
    Dim V As Long
    On Error GoTo ErrHandler
    V = Val(textbox(Index))
    Select Case Index
        Case 0
            If V > 0 And V <= 5000 Then
                ScanRate = V
                UpdateForm
            Else
                Beep
                Cancel = True
            End If
        Case 1
            If V > 0 And V <= 5000 Then
                ReadInterval = V
                UpdateForm
            Else
                Beep
                Cancel = True
            End If
        Case 2
            If V >= 0 And V <= 100 Then
                Debounce = V
                UpdateForm
            Else
                Beep
                Cancel = True
            End If
        Case 3
            If V >= 0 And V <= 100 Then
                AlarmSet = V
                UpdateForm
            Else
                Beep
                Cancel = True
            End If
        Case 4
            If V > 0 And V < 65000 Then
                Port = V
                UpdateForm
            Else
                Beep
                Cancel = True
            End If
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "Textbox_Validate", Err.Description
    Resume ErrExit
End Sub

Private Sub UpdateForm()
    Dim F As Long
    On Error GoTo ErrHandler
    For F = 0 To FldCount - 1
        Select Case F
            Case 0
                textbox(F) = ScanRate
            Case 1
                textbox(F) = ReadInterval
            Case 2
                textbox(F) = Debounce
            Case 3
                textbox(F) = AlarmSet
            Case 4
                textbox(F) = Port
        End Select
    Next F
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSettings", "UpdateForm", Err.Description
    Resume ErrExit
End Sub

