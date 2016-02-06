VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "JwldButn2b.ocx"
Begin VB.Form frmRestore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextBox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin JwldButn2b.JeweledButton cmdOK 
      Height          =   390
      Left            =   3360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   1140
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Close"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton cmdCancel 
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1140
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Cancel"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1185
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "New name for file"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReSetting As Boolean
Public LoadOk As Boolean
Dim CancelOpen As Boolean
Private Sub cmdCancel_Click()
    ResetStatus
End Sub
Private Sub cmdOK_LostFocus()
        StatusBar1.SimpleText = ""
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    Select Case KeyCode
        Case 38
            'up arrow
            SendKeys ("+{tab}")
            KeyCode = 0
        Case 40
            'down arrow
            SendKeys ("{tab}")
            KeyCode = 0
    End Select
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmSaveAs", "Form_KeyDown", Err.Description
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
    LoadOk = False
    ResetStatus
    LoadOk = True
ErrExit:
End Sub
Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub
Private Sub ResetStatus()
    Dim CK As Long
    On Error GoTo ErrHandler
    ReSetting = True
    For CK = 0 To 0
        TextBox(CK) = ""
    Next CK
    cmdCancel.Enabled = False
    cmdOK.Caption = "Close"
    ReSetting = False
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmSaveAs", "ResetStatus", Err.Description
    Resume ErrExit
End Sub
Private Sub textBox_Change(Index As Integer)
    UpdateStatus
End Sub
Private Sub textbox_GotFocus(Index As Integer)
    TextBox(Index).SelStart = 0
    TextBox(Index).SelLength = Len(TextBox(Index).Text)
End Sub
Private Sub UpdateStatus()
    Dim CK As Long
    Dim Changed As Boolean
    On Error GoTo ErrHandler
    If Not ReSetting Then
        For CK = 0 To 0
            If TextBox(CK) <> "" Then
                Changed = True
                Exit For
            End If
        Next CK
        If Changed Then
            cmdCancel.Enabled = True
            cmdOK.Caption = "Save"
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmSaveAs", "UpdateStatus", Err.Description
    Resume ErrExit
End Sub
Private Sub GetFile()
    On Error GoTo ErrHandler
    With frmStart.ComDial
        .FileName = ""  'prevents using previous directory
        .CancelError = True 'raises error cdlCancel if user presses the cancel button
        .InitDir = AD.AppData("BackupLocation")
        .Filter = "Database files|*.mdb"
        .Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
        .DialogTitle = "Open"
        .ShowOpen
        DbName = .FileName
    End With
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case cdlCancel
            CancelOpen = True
        Case Else
            DisplayError Err.Number, "FileUtils", "OpenExistingDb", Err.Description
    End Select
    Resume ErrExit
End Sub
