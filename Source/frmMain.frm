VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Temperature Monitor"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComDial 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   960
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox tbEvents 
      Height          =   7335
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton butServer 
      Caption         =   "Enable Server"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSaveAs 
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
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "File Information"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSensors 
      Caption         =   "Sensors"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuMonHelp 
         Caption         =   "BinMonitor Help"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Log"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const mConMaxConnections As Integer = 3
Const PortNum As Long = 1500
Dim ServerEnabled As Boolean

Private Sub butServer_Click()
    On Error GoTo ErrHandler
    ServerEnabled = Not ServerEnabled
    CheckServerStatus
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "butServer_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub CheckClients()
'---------------------------------------------------------------------------------------
' Procedure : CheckClients
' Author    : David
' Date      : 1/1/2012
' Purpose   : check if all clients are still connected
'---------------------------------------------------------------------------------------
'
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    For Each Sck In sckServer
        If Sck.State = sckClosed Or Sck.State = sckError Or Sck.State = sckClosing Then sckServer_Close (Sck.Index)
    Next
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "CheckClients", Err.Description
    Resume ErrExit
End Sub

Private Sub CheckServerStatus()
    On Error GoTo ErrHandler
    If ServerEnabled Then
        butServer.Caption = "Close Server"
        If sckServer(0).State = sckClosed Then
            'initialize the port on which to listen
            sckServer(0).LocalPort = PortNum
            'Start listening for a client connection request
            sckServer(0).Listen
            EventMessage "Server Listening ..."
        End If
    Else
        butServer.Caption = "Enable Server"
        If sckServer(0).State <> sckClosed Then sckServer(0).Close
        EventMessage "Server closed."
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 10048
            'port is in use
        Case Else
            AD.DisplayError Err.Number, "frmMain", "CheckServerStatus", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub Command1_Click()
    Dim S As String
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    If ServerEnabled Then
        CheckClients
        S = Text1
        'sent to all
        For Each Sck In sckServer
            If Sck.Index > 0 Then
                Sck.SendData S
                DoEvents
            End If
        Next Sck
    End If
    On Error GoTo 0
ErrExit:
    DoEvents
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "Command1_Click", Err.Description
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
    AD.DisplayError Err.Number, "frmMain", "EventMessage", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Activate()
    ConnectDatabase
    Me.Caption = "BinMonitor"
    If Prog.IsValid Then
        Me.Caption = Me.Caption & " - [" & Prog.DatabaseName & "]"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    AD.LoadFormData Me
    CheckServerStatus
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "Form_Load", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub

Private Sub mnuAbout_Click()
    Dim Mes As String
    On Error GoTo ErrHandler
    Mes = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Mes = Mes & vbCr
    Mes = Mes & VersionDate
    Call MsgBox(Mes, vbInformation, App.Title)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "mnuAbout_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub ProcessData(ND As String)
    Dim DataType As PacketData
    Dim Newval As String
    Dim ID As Long
    Dim Pkts() As String
    Dim Count As Long
    Dim C As Long
    Dim Parts() As String
    On Error GoTo ErrHandler
    Pkts = Split(ND, BeginPacket)
    Count = UBound(Pkts)
    For C = 0 To Count
        Parts = Split(Pkts(C), "|")
        'check for correctly formed packet
        If UBound(Parts) = 2 Then
            DataType = Val(Parts(0))
            Newval = Parts(1)
            ID = Val(Parts(2))
            Select Case DataType
                Case PacketData.pdTemp
                
'                Case PacketData.pdConnection
'                    If NewVal = "Initialize" Then
'                        'do client setup
'                        SetupClientControls
'                    End If
            End Select
        End If
    Next C
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "ProcessData", Err.Description
     Resume ErrExit
End Sub

Private Sub mnuNew_Click()
    frmNewDatabase.Show vbModal
End Sub

Private Sub sckServer_Close(Index As Integer)
    On Error GoTo ErrHandler
    ' If not the listening socket then ...
    If Index > 0 Then
'        SendPacket pdStatus, Str(AlarmStatusType.astOff), , Index
        ' Make sure the connection is closed
        If sckServer(Index).State <> sckClosed Then sckServer(Index).Close
        ' Unload the control
        Unload sckServer(Index)
        EventMessage "Socket " & CStr(Index) & " - Connection closed"
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "sckServer_Close", Err.Description
    Resume ErrExit
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    ' Accept the connection on the next socket control in the control array
    Dim ID As Integer
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    ' Ignore request if not the server socket.
    If Index = 0 Then
        CheckClients
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
        If ID <= mConMaxConnections Then
            ' Load new control and accept connection
            Load sckServer(ID)
            sckServer(ID).LocalPort = 0
            sckServer(ID).Accept requestID
            ' Indicate status
            EventMessage "Socket " & CStr(ID) & " accepted a connection"
'            SetupClientControls ID
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "sckServer_ConnectionRequest", Err.Description
    Resume ErrExit
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    On Error GoTo ErrHandler
    'socket 0 is only for making connections
    If Index > 0 Then
        sckServer(Index).GetData strData, vbString
        EventMessage strData
        ProcessData strData
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "sckServer_DataArrival", Err.Description
     Resume ErrExit
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AD.SaveToLog Description, "frmMain", "sckClient_Error" & "(" & Index & ")", CLng(Number), adError
    CancelDisplay = True
End Sub

Private Sub SendPacket(DataType As PacketData, Newval As String, Optional ID As _
    Long = 0, Optional Client As Integer = -1)
    Dim S As String
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    If ServerEnabled Then
        CheckClients
        S = BeginPacket & DataType & "|" & Newval & "|" & ID
        If Client = -1 Then
            'sent to all
            For Each Sck In sckServer
                If Sck.Index > 0 Then
                    Sck.SendData S
                    DoEvents
                End If
            Next Sck
        Else
            'send to specified
            If Client > 0 And Client <= sckServer.Count Then
                sckServer(Client).SendData S
            End If
        End If
    End If
    On Error GoTo 0
ErrExit:
    DoEvents
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "SendPacket", Err.Description
     Resume ErrExit
End Sub

