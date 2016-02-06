VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3CC26050-93EE-4497-B39F-C5F1BA7EAA84}#2.1#0"; "dcBinMaps.ocx"
Begin VB.Form frmStart 
   BackColor       =   &H8000000C&
   Caption         =   "TemperatureMonitor"
   ClientHeight    =   7575
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11655
   HelpContextID   =   10
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11655
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enable Network"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Record data"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enable Notifications"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enable Maps"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "File Information"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":158A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":2B24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog ComDial 
      Left            =   5160
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin dcBinMaps.BinMaps BinMaps1 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   428
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      MapType         =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7920
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MenuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu MnuCompact 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mnuServer 
         Caption         =   "Enable Server"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "Enable Recording"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuNotify 
         Caption         =   "Enable Notification"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuUseMaps 
         Caption         =   "Use Bin Maps"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "File Information"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSensors 
      Caption         =   "Sensors"
   End
   Begin VB.Menu mnuClients 
      Caption         =   "Clients"
   End
   Begin VB.Menu mnuEditBins 
      Caption         =   "Bins"
   End
   Begin VB.Menu mnuMaps 
      Caption         =   "Maps"
   End
   Begin VB.Menu mnuBins 
      Caption         =   "Bin Report"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "TemperatureMonitor Help"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Log"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxLinks As Integer = 5
Const PortNum As Long = 1600

Private Type Requests
    Start As Date
    Sckt As Integer
    ReplyReceived As Boolean
End Type

Dim Request() As Requests
Private MapEdited As Boolean
Private BinMapErrorCount As Long
Dim StatusDelay As Long
Dim RecordStart As Date
Dim Status As String
Dim SocketCount As Long '# of connected sockets
Dim AlarmStart As Date

Public LoadOK As Boolean

Private Sub BinMaps1_BinRepositioned()
    MapEdited = True
End Sub

Private Sub BinMaps1_ToopTip(TipText As String)
    ShowStatus TipText
End Sub

Private Sub CheckAlarms()
'---------------------------------------------------------------------------------------
' Procedure : CheckAlarms
' Author    : XPMUser
' Date      : 1/17/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Dim Col As clsSensors
    Dim Obj As clsSensor
    Dim Tmp As Currency
    On Error GoTo ErrHandler
    If DateDiff("h", AlarmStart, Now) > Prog.AlarmInterval Then
        AlarmStart = Now
        Set Col = New clsSensors
        Col.Load
        For Each Obj In Col
            Obj.BeginEdit
            Obj.Status = ssEnabled
            'daily temp
            Tmp = Obj.TempLastDay
            Obj.DailyTemp = Tmp
            If Tmp > Obj.MaxTemp Then
                Obj.Status = ssAlarmCondition
                AD.SaveStatus "Bin # " & Obj.BinDescription & ", Sensor # " & Obj.Number & " temperature over maximum."
                AD.SaveToLog "Bin # " & Obj.BinDescription & ", Sensor # " & Obj.Number & " temperature over maximum.", , , , adshort
            End If
            'trend temp
            Tmp = Obj.Trend(Prog.TrendTime)
            Obj.TrendTemp = Tmp
            If Tmp > Prog.TrendMax Then
                Obj.Status = ssAlarmCondition
                AD.SaveStatus "Bin # " & Obj.BinDescription & ", Sensor # " & Obj.Number & " trend temperature over maximum."
                AD.SaveToLog "Bin # " & Obj.BinDescription & ", Sensor # " & Obj.Number & " trend temperature over maximum.", , , , adshort
            End If
            Obj.ApplyEdit
        Next
    End If
    Set Col = Nothing
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "CheckAlarms", Err.Description
     Resume ErrExit
End Sub

Private Sub CheckDeadSockets()
'---------------------------------------------------------------------------------------
' Procedure : CheckDeadSockets
' Author    : David
' Date      : 1/1/2012
' Purpose   : check if any sockets are not connected
'---------------------------------------------------------------------------------------
'
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    For Each Sck In sckServer
        If Sck.State = sckClosed Or Sck.State = sckError Or Sck.State = sckClosing Then
            sckServer_Close (Sck.Index)
        End If
    Next
    CountSockets
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "CheckDeadSockets", Err.Description
    Resume ErrExit
End Sub

Private Sub CheckServerStatus()
    On Error GoTo ErrHandler
    If mnuServer.Checked Then
        If sckServer(0).State = sckClosed Then
            'initialize the port on which to listen
            sckServer(0).LocalPort = PortNum
            'Start listening for a client connection request
            sckServer(0).Listen
            ShowStatus "Server Listening ..."
        End If
    Else
        CloseSockets
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 10048
            'port is in use
        Case Else
            AD.DisplayError Err.Number, "frmStart", "CheckServerStatus", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Function ClientID(Sckt As Integer) As Long
'---------------------------------------------------------------------------------------
' Procedure : ClientID
' Author    : XPMUser
' Date      : 1/30/2016
' Purpose   : return client ID for socket, request update if needed
'---------------------------------------------------------------------------------------
'
    Dim Obj As clsClient
    On Error GoTo ErrHandler
    Set Obj = New clsClient
    Obj.Load , , Sckt
    If Obj.IsNew Then
        ClientID = 0
    Else
        ClientID = Obj.ID
    End If
    If Obj.IsNew Or Obj.SocketID = 0 Then
        'request update
        RequestClientMac Sckt
    End If
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "ClientID", Err.Description
     Resume ErrExit
End Function

Private Sub CloseSockets()
    Dim S As Winsock
    On Error GoTo ErrHandler
    For Each S In sckServer
        If S.State <> sckClosed Then S.Close
    Next S
    Set S = Nothing
    ShowStatus "Server closed."
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "CloseSockets", Err.Description
     Resume ErrExit
End Sub

Private Sub CountSockets()
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    SocketCount = 0
    For Each Sck In sckServer
        If Sck.Index <> 0 And Sck.State <> sckClosed And Sck.State <> sckError And Sck.State <> sckClosing Then
            SocketCount = SocketCount + 1
        End If
    Next
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "CountSockets", Err.Description
     Resume ErrExit
End Sub

Private Sub DBinit()
'---------------------------------------------------------------------------------------
' Procedure : DBinit
' Author    : David
' Date      : 1/31/2016
' Purpose   : erase all sockets from database, erase client ID's from sensors
'---------------------------------------------------------------------------------------
'
    Dim Obj As clsClient
    Dim Col As clsClients
    Dim objSensor As clsSensor
    Dim ColSensors As clsSensors
    On Error GoTo ErrHandler
    Set Col = New clsClients
    Col.Load
    For Each Obj In Col
        With Obj
            .BeginEdit
            .SocketID = 0
            .ApplyEdit
        End With
    Next
    Set ColSensors = New clsSensors
    ColSensors.Load
    For Each objSensor In ColSensors
        With objSensor
            .BeginEdit
            .ClientID = 0
            .ApplyEdit
        End With
    Next
    Set Obj = Nothing
    Set Col = Nothing
    Set objSensor = Nothing
    Set ColSensors = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "DBinit", Err.Description
     Resume ErrExit
End Sub

Private Sub DBremoveSocket(Sckt As Integer)
'---------------------------------------------------------------------------------------
' Procedure : DBremoveSocket
' Author    : David
' Date      : 1/31/2016
' Purpose   : remove socket ID from database when socket closes
'---------------------------------------------------------------------------------------
'
    Dim Obj As clsClient
    On Error GoTo ErrHandler
    Set Obj = New clsClient
    Obj.Load , , Sckt
    If Not Obj.IsNew Then
        'erase socket
        With Obj
            .BeginEdit
            .SocketID = 0
            .ApplyEdit
        End With
    End If
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "DBremoveSocket", Err.Description
     Resume ErrExit
End Sub

Private Sub DisableBinMaps()
    BinMapErrorCount = BinMapErrorCount + 1
    If BinMapErrorCount - 1 = 0 Then
        mnuUseMaps.Checked = False
        BinMaps1.Visible = False
    Else
        '2nd error, stop program
        Call MsgBox("TemperatureMonitor can not start.", vbInformation Or vbSystemModal, App.Title)
        Unload Me
    End If
End Sub

Public Sub EnableMenus(IsValid As Boolean)
'---------------------------------------------------------------------------------------
' Procedure : EnableMenus
' Author    : David
' Date      : 07/Feb/2010
' Purpose   :
' Feb 12, 2012 updated for version
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Me.Caption = "TemperatureMonitor"
    If IsValid Then
        Me.Caption = Me.Caption & " - [" & Prog.DatabaseName & "]"
    End If
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "EnableMenus", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Activate()
    If DBconnected Then
        UpdateMaps
        Timer1.Enabled = True
        mnuServer.Checked = True
        UpdateButtons
        CheckServerStatus
    End If
    'check frequently
    EnableMenus Prog.IsValid
End Sub

Private Sub Form_Load()
    Dim UM As String
    AD.LoadFormData Me
    ReDim Request(0)
    ConnectDatabase
    If DBconnected Then
        RecordStart = DateAdd("n", -3 * Prog.RecordInterval, Now)  'so a reading is taken immediately
        DBinit
    End If
    mnuServer.Checked = False
    CheckDeadSockets
    UM = AD.AppData("ShowMaps")
    If UM = "" Then
        mnuUseMaps.Checked = False
    Else
        mnuUseMaps.Checked = CBool(UM)
    End If
    CheckServerStatus
    UpdateButtons
    LoadOK = True
End Sub

Private Sub Form_Resize()
    BinMaps1.Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    Dim C As Long
    Dim FF As Form
    'unload all forms
    AD.SaveFormData Me
    AD.AppData("ShowMaps") = mnuUseMaps.Checked
    SaveMaps BinMaps1
    For Each FF In Forms
        C = C + 1
        If C > 100 Then Exit For
        Unload FF
    Next FF
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "Form_Unload", Err.Description
    Resume ErrExit
End Sub

Private Sub MenuSaveAs_Click()
    ShowForm ("frmSaveAs")
End Sub

Private Sub mnuBackup_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuBackup_Click
' Author    : David
' Date      : 3/6/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Dim BK As String
    On Error GoTo ErrHandler
    BK = AD.Folders(App_Folders_Backup)
    If CopyCurrentFile(BK) Then
        ShowStatus "Data backed-up."
    Else
        ShowStatus "Data could not be backed-up."
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "mnuBackup_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub mnuBins_Click()
    frmBinReport.Show vbModal
End Sub

Private Sub mnuClients_Click()
    On Error GoTo ErrHandler
    frmClients.Show vbModal
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuClients_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub MnuCompact_Click()
    Dim CompactWorked As Boolean
    'compact & repair database
    On Error GoTo ErrHandler
    CompactWorked = Prog.CompactDatabase
ErrExit:
    If CompactWorked Then
        ShowStatus "Database Compacted."
    Else
        Beep
        ShowStatus "Failed to Compact database."
        DBconnected = False
        Form_Activate
    End If
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 3356
            'locked, do nothing
        Case Else
            AD.DisplayError Err.Number, "frmStart", "MnuCompact_Click", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub mnuEditBins_Click()
    frmStorage.Show vbModal
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuFileOpen_Click
' Author    : David
' Date      : 12/5/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim DBname As String
    On Error GoTo ErrHandler
    CloseSockets
    With frmStart.ComDial
        .FileName = ""  'prevents using previous directory
'                    .CancelError = True 'raises error cdlCancel if user presses the cancel button
        .InitDir = AD.Folders(App_Folders_Database)
        .Filter = "Database files|*.mdb"
        .Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
        .DialogTitle = "Open"
        .ShowOpen
        DBname = .FileName
    End With
    If ConnectDatabase(DBname) Then
        UpdateMaps
        CheckServerStatus
    Else
        ShowStatus "File not opened."
        ConnectDatabase 'to original file
    End If
    EnableMenus Prog.IsValid
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "mnuFileOpen_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub mnuHelpAbout_Click()
    On Error GoTo ErrHandler
    frmAbout.Show vbModal
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuHelpAbout_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuHelpHelp_Click()
'    SendKeys "{F1}"
End Sub

Private Sub mnuInfo_Click()
    frmInfo.Show vbModal
End Sub

Private Sub mnuLog_Click()
    On Error GoTo ErrExit
    frmLog.Show vbModal
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "mnuLog_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub mnuMaps_Click()
    On Error GoTo ErrHandler
    frmBinMaps.Show vbModal
    UpdateMaps
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuMaps_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuNew_Click()
    ShowForm "frmNewDatabase"
    UpdateMaps
End Sub

Private Sub MnuNotify_Click()
    MnuNotify.Checked = Not MnuNotify.Checked
    UpdateButtons
End Sub

Private Sub mnuOptions_Click()
    On Error GoTo ErrHandler
    frmOptions.Show vbModal
    AlarmStart = 0
    CheckAlarms
    UpdateMaps
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuOptions_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuPassword_Click()
    frmChangePassword.Show vbModal
End Sub

Private Sub mnuRecord_Click()
    mnuRecord.Checked = Not mnuRecord.Checked
    UpdateButtons
End Sub

Private Sub mnuRestore_Click()
'---------------------------------------------------------------------------------------
' Procedure : mnuRestore_Click
' Author    : David
' Date      : 3/5/2012
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Dim FSO As FileSystemObject
    Dim FL As File
    Dim DT As String
    Dim Tm As String
    Dim P As Long
    Dim Src As String
    Dim Des As String
    Dim DocSrc As String
    Dim IsCurrent As Boolean
    On Error GoTo ErrHandler
    With frmStart.ComDial
        .FileName = ""  'prevents using previous directory
        .CancelError = True 'raises error cdlCancel if user presses the cancel button
        .InitDir = AD.Folders(App_Folders_Backup)
        .Filter = "Database files|*.mdb"
        .Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
        .DialogTitle = "File Restore"
        .ShowOpen
        Src = .FileName
    End With
    If Src <> "" Then
        'remove the .mdb
        DocSrc = Left$(Src, Len(Src) - 4)
        Des = AD.Folders(App_Folders_Database) & "\" & DatabaseName(Src)
        'check if restoring current database
        IsCurrent = (LCase(DatabaseName(Src)) = LCase(Prog.DatabaseName))
        Set FSO = New FileSystemObject
        Set FL = FSO.GetFile(Src)
        DT = FL.DateLastModified
        P = InStr(DT, " ")
        Tm = Right$(DT, Len(DT) - P)
        DT = Left$(DT, P - 1)
        DT = Format(DT, "medium date")
        Select Case MsgBox("Do you want to restore this database that was backed-up on " & DT & " " & Tm & " ?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, App.Title)
            Case vbYes
                If IsCurrent Then
                    'close database
                    Set Prog = Nothing
                    Set Prog = New clsMain
                End If
                'copy database
                FSO.CopyFile Src, Des & ".mdb"
                'copy document folder
                If FSO.FolderExists(DocSrc) Then
                    FSO.CopyFolder DocSrc, Des
                End If
                If IsCurrent Then
                    'reopen database
                    ConnectDatabase Des & ".mdb"
                End If
                ShowStatus ("Database restored.")
            Case vbNo
                ShowStatus "Database not restored."
        End Select
    End If
    On Error GoTo 0
ErrExit:
    Set FSO = Nothing
    Set FL = Nothing
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case cdlCancel
'            CancelOpen = True
        Case Else
            AD.DisplayError Err.Number, "frmStart", "mnuRestore_Click", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub mnuSensors_Click()
    On Error GoTo ErrHandler
    frmSensors.Show vbModal
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuSensors_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuServer_Click()
    On Error GoTo ErrHandler
    mnuServer.Checked = Not mnuServer.Checked
    UpdateButtons
    CheckServerStatus
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "mnuServer_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuUseMaps_Click()
    mnuUseMaps.Checked = Not mnuUseMaps.Checked
    UpdateButtons
    UpdateMaps
End Sub

Private Sub PassMessage(NewMessage As String, frmName As String)
    Dim F As Form
    Dim Found As Boolean
    On Error GoTo ErrHandler
    For Each F In Forms
        If F.Name = frmName Then
            Found = True
            Exit For
        End If
    Next
    If Found Then
        Select Case LCase(frmName)
            Case "frmsensors"
                frmSensors.EventMessage NewMessage
            Case "frmclients"
                frmClients.EventMessage NewMessage
        End Select
    End If
    Set F = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "PassMessage", Err.Description
     Resume ErrExit
End Sub

Private Sub ProcessData(ND As String, Sckt As Integer)
    Dim DataType As PacketData
    Dim NewVal As String
    Dim RomCode As String
    Dim SensorID As Long
    Dim Pkts() As String
    Dim Count As Long
    Dim C As Long
    Dim Parts() As String
    Dim Mes As String
    On Error GoTo ErrHandler
    Pkts = Split(ND, BeginPacket)
    Count = UBound(Pkts)
    For C = 0 To Count
        Parts = Split(Pkts(C), "|")
        'check for correctly formed packet
        If UBound(Parts) = 3 Then
            DataType = Val(Parts(0))
            NewVal = Parts(1)
            RomCode = TR(Parts(2))
            'process packets
            Select Case DataType
                Case PacketData.pdGetTemperatures
                    Mes = "Sensor " & SensorName(RomCode, Sckt, SensorID) & " Temperature = " & NewVal
                    ShowStatus Mes, "frmSensors"
                    If mnuRecord.Checked Then SaveTemp SensorID, NewVal
                Case PacketData.pdGetBinNum
                    Mes = "Sensor " & SensorName(RomCode, Sckt) & " Bin Number = " & NewVal
                    ShowStatus Mes, "frmSensors"
                Case PacketData.pdGetSenNum
                    Mes = "Sensor " & SensorName(RomCode, Sckt) & " Sensor Number = " & NewVal
                    ShowStatus Mes, "frmSensors"
                Case PacketData.pdGetSignalStrength
                    ShowSignal NewVal, Sckt
                    ClientID Sckt   'to check if sckt is associated with a client
                Case PacketData.pdGetClientMac
                    Mes = "Client Mac = " & NewVal
                    SaveSocketClient NewVal, Sckt
            End Select
        End If
    Next C
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "ProcessData", Err.Description
     Resume ErrExit
End Sub

Private Sub RecordData()
    On Error GoTo ErrHandler
    If DateDiff("n", RecordStart, Now) > Prog.RecordInterval Then
        SendPacket pdGetTemperatures
        RecordStart = Now
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "RecordData", Err.Description
     Resume ErrExit
End Sub

Private Sub RequestClientMac(SocketID As Integer)
'---------------------------------------------------------------------------------------
' Procedure : RequestClientMac
' Author    : XPMUser
' Date      : 1/4/2016
' Purpose   : send a request on socket to get client mac address
'---------------------------------------------------------------------------------------
'
    Dim R As Long
    Dim Found As Boolean
    For R = 1 To UBound(Request)
        If Request(R).Sckt = SocketID Then
            'existing Request
            Found = True
            'only send every 60 seconds
            If DateDiff("s", Request(R).Start, Now) > 60 Then
                Request(R).Start = Now
                SendPacket pdGetClientMac, , , SocketID
            End If
            Exit For
        End If
    Next R
    If Not Found Then
        'new request
        R = UBound(Request) + 1
        ReDim Preserve Request(R)
        Request(R).Sckt = SocketID
        Request(R).Start = Now
        SendPacket pdGetClientMac, , , SocketID
    End If
End Sub

Private Sub SaveSocketClient(Mac As String, Sckt As Integer)
    Dim Obj As clsClient
    On Error GoTo ErrHandler
    Set Obj = New clsClient
    Obj.Load , Mac
    If Obj.IsNew Or Obj.SocketID = 0 Then
        'update
        With Obj
            .BeginEdit
            .SocketID = Sckt
            .Mac = Mac
            .ApplyEdit
        End With
    End If
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "SaveSocketClient", Err.Description
     Resume ErrExit
End Sub

Private Sub SaveTemp(SensorID As Long, NewVal As String)
    Dim Obj As clsRecord
    On Error GoTo ErrHandler
    Set Obj = New clsRecord
    With Obj
        .BeginEdit
        .SensorID = SensorID
        .Temperature = NewVal
        .recDate = Now
        .ApplyEdit
    End With
ErrExit:
    Set Obj = Nothing
    Exit Sub
ErrHandler:
    AD.SaveToLog "Could not record data.", "frmStart", "SaveTemp", Err.Number, adshort
    Resume ErrExit
End Sub

Private Sub sckServer_Close(Index As Integer)
    On Error GoTo ErrHandler
    ' If not the listening socket then ...
    If Index > 0 Then
        DBremoveSocket Index
        ' Make sure the connection is closed
        If sckServer(Index).State <> sckClosed Then sckServer(Index).Close
        ' Unload the control
        Unload sckServer(Index)
        ShowStatus "Socket " & CStr(Index) & " - Connection closed"
    End If
    CountSockets
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "sckServer_Close", Err.Description
    Resume ErrExit
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    ' Accept the connection on the next socket control in the control array
    Dim ID As Integer
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    ' Ignore request if not the server socket.
    If Index = 0 Then
        CheckDeadSockets
        ' Look for an available socket control index
        ID = 1
        For Each Sck In sckServer
            ' Skip index 0 as this is the listening socket
            ' this accounts for sockets that may have been
            ' unloaded. Instead of sockets 0,1,2,3 in the collection
            ' it may be 0,1,3. The next socket selected would be 2.
            If Sck.Index > 0 Then
                If ID < Sck.Index Then Exit For
                ID = ID + 1
            End If
        Next Sck
        ' check for maximum # of controls
        If ID <= MaxLinks Then
            ' Load new control and accept connection
            Load sckServer(ID)
            sckServer(ID).LocalPort = 0
            sckServer(ID).Accept requestID
            ' Indicate status
            ShowStatus "Socket " & CStr(ID) & " accepted a connection"
            'check for new clients
            RequestClientMac ID
        End If
    End If
    CountSockets
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmStart", "sckServer_ConnectionRequest", Err.Description
    Resume ErrExit
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    On Error GoTo ErrHandler
    'socket 0 is only for making connections
    If Index > 0 Then
        sckServer(Index).GetData strData, vbString
        ProcessData strData, Index
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "sckServer_DataArrival", Err.Description
     Resume ErrExit
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AD.SaveToLog Description, "frmStart", "sckClient_Error" & "(" & Index & ")", CLng(Number), adError
    CancelDisplay = True
End Sub

Public Sub SendPacket(DataType As PacketData, Optional NewVal As String = "", _
    Optional SensorID As Long = 0, Optional SocketID As Integer = -1)
    Dim S As String
    Dim Sck As Winsock
    Dim RomCode As String
    On Error GoTo ErrHandler
    If SensorID <> 0 Then SocketID = SensorSocket(SensorID, RomCode)
    If mnuServer.Checked Then
        CheckDeadSockets
        S = BeginPacket & DataType & "|" & NewVal & "|" & RomCode & "|"
        If SocketID = -1 Then
            'sent to all
            For Each Sck In sckServer
                If Sck.Index > 0 Then
                    Sck.SendData S
                    DoEvents
                End If
            Next Sck
        Else
            'send to specified
            If SocketID > 0 And SocketID <= sckServer.Count Then
                sckServer(SocketID).SendData S
            End If
        End If
    End If
    On Error GoTo 0
ErrExit:
    DoEvents
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "SendPacket", Err.Description
     Resume ErrExit
End Sub

Private Sub SensorChanged()
'---------------------------------------------------------------------------------------
' Procedure : SensorChanged
' Author    : David
' Date      : 2/1/2016
' Purpose   : update grid on frmSensors if loaded
'---------------------------------------------------------------------------------------
'
    Dim F As Form
    Dim Found As Boolean
    On Error GoTo ErrHandler
    For Each F In Forms
        If F.Name = "frmSensors" Then
            Found = True
            Exit For
        End If
    Next
    If Found Then frmSensors.SensorsChanged
    Set F = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "SensorChanged", Err.Description
     Resume ErrExit
End Sub

Private Function SensorName(RomCode As String, Sckt As Integer, Optional SensorID As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : SensorName
' Author    : XPMUser
' Date      : 1/30/2016
' Purpose   : returns SensorName and SensorID
'---------------------------------------------------------------------------------------
'
    Dim Obj As clsSensor
    On Error GoTo ErrHandler
    Set Obj = New clsSensor
    Obj.Load 0, RomCode
    If Obj.IsNew Then
        'new sensor
        With Obj
            .BeginEdit
            .RomCode = RomCode
            .ClientID = ClientID(Sckt)
            .ApplyEdit
        End With
        SensorName = "' " & RomCode & " ', "
        SensorChanged
    Else
        'existing sensor
        SensorName = Obj.Number & ",  Bin " & Obj.BinDescription & ", "
        If Obj.ClientID = 0 Then
            'update clientID
            With Obj
                .BeginEdit
                .ClientID = ClientID(Sckt)
                .ApplyEdit
            End With
            SensorChanged
        End If
    End If
    SensorID = Obj.ID
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "SensorName", Err.Description
     Resume ErrExit
End Function

Private Function SensorSocket(SensorID As Long, RomCode As String) As Integer
'---------------------------------------------------------------------------------------
' Procedure : SensorSocket
' Author    : XPMUser
' Date      : 1/12/2016
' Purpose   : return # of socket sensor is connected to
'             returns -1 if not found
'             also returns sensor Rom Code
'---------------------------------------------------------------------------------------
'
    Dim Obj As clsSensor
    Dim Sck As Winsock
    Dim ObjClient As clsClient
    On Error GoTo ErrHandler
    SensorSocket = -1
    Set Obj = New clsSensor
    Obj.Load SensorID
    If Obj.IsValid Then
        RomCode = Obj.RomCode
        Set ObjClient = New clsClient
        ObjClient.Load Obj.ClientID
        SensorSocket = ObjClient.SocketID
    End If
    Set Obj = Nothing
    Set ObjClient = Nothing
    Set Sck = Nothing
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "SensorSocket", Err.Description
     Resume ErrExit
End Function

Private Sub ShowSignal(NewVal As String, Sckt As Integer)
    Dim Obj As clsClient
    Dim Mes As String
    On Error GoTo ErrHandler
    Set Obj = New clsClient
    Obj.Load , , Sckt
    If Obj.IsValid Then
        Mes = "Client " & Obj.Description & " signal strength is " & NewVal
        ShowStatus Mes, "frmClients"
    Else
        Mes = "Unknown Client signal strength is " & NewVal
        ShowStatus Mes, "frmClients"
    End If
    Set Obj = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "ShowSignal", Err.Description
     Resume ErrExit
End Sub

Private Sub ShowStatus(Optional Message As String, Optional frmName As String)
    'show on this form
    If Message = "" Then
        Status = ""
    Else
        If Status = "" Then
            Status = Message
        Else
            Status = Status & " / " & Message
        End If
    End If
    StatusBar1.SimpleText = "Clients Connected = " & SocketCount & " : " & Status
    'show on other forms
    If frmName <> "" Then PassMessage Message, frmName
    StatusDelay = 4 '4 loops of 2 seconds = 8 seconds
End Sub

Private Sub Timer1_Timer()
    StatusDelay = StatusDelay - 1
    If StatusDelay < 1 Then ShowStatus
    SendPacket PacketData.pdHeartBeat
    If mnuRecord.Checked Then RecordData
    If MnuNotify.Checked Then CheckAlarms
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrHandler
    Select Case Button.Index
        Case 2
            'server
            mnuServer_Click
        Case 3
            'record
            mnuRecord_Click
        Case 4
            'notifications
            MnuNotify_Click
        Case 5
            'maps
            mnuUseMaps_Click
        Case 6
            'info
            mnuInfo_Click
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "Toolbar1_ButtonClick", Err.Description
     Resume ErrExit
End Sub

Private Sub UpdateButtons()
    On Error GoTo ErrHandler
    With Toolbar1
        If mnuServer.Checked Then
            .Buttons(2).Value = tbrPressed
        Else
            .Buttons(2).Value = tbrUnpressed
        End If
        If mnuRecord.Checked Then
            .Buttons(3).Value = tbrPressed
        Else
            .Buttons(3).Value = tbrUnpressed
        End If
        If MnuNotify.Checked Then
            .Buttons(4).Value = tbrPressed
        Else
            .Buttons(4).Value = tbrUnpressed
        End If
        If mnuUseMaps.Checked Then
            .Buttons(5).Value = tbrPressed
        Else
            .Buttons(5).Value = tbrUnpressed
        End If
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmStart", "UpdateButtons", Err.Description
     Resume ErrExit
End Sub

Private Sub UpdateMaps()
    On Error GoTo ErrHandler
    ConnectMaps BinMaps1, mnuUseMaps.Checked
    BinMapErrorCount = 0
ErrExit:
    Exit Sub
ErrHandler:
    'stop using maps if there is an error
    AD.DisplayError Err.Number, "frmStart", "UpdateMaps", Err.Description
    Call MsgBox("Maps will be disabled.", vbInformation Or vbSystemModal, App.Title)
    DisableBinMaps
    Resume ErrExit
End Sub

