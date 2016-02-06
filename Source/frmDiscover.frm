VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{D9794188-4A19-4061-A8C1-BCC9E392E6DB}#3.1#0"; "dcCombo.ocx"
Object = "{6E1B7648-872D-4B19-96AD-0555B4151387}#14.0#0"; "dcList.ocx"
Begin VB.Form frmDiscover 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discover New Sensors"
   ClientHeight    =   11655
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11655
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin dcList.dcListControl List 
      Height          =   2895
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5106
      SortCol         =   0
      RowsToDisplay   =   8
      MinRowsToDisplay=   8
      HideSearch      =   0   'False
      GridCaption     =   "Sensors"
   End
   Begin dcCombo.dcComboControl Combo 
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      HideAdd         =   -1  'True
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtSensor 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1950
      Width           =   975
   End
   Begin VB.TextBox txtBin 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton butSetBin 
      Caption         =   "Set"
      Height          =   315
      Left            =   4440
      TabIndex        =   7
      Top             =   2370
      Width           =   1335
   End
   Begin VB.CommandButton butSetSensor 
      Caption         =   "Set"
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton butGetBin 
      Caption         =   "Get"
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2370
      Width           =   1335
   End
   Begin VB.CommandButton butGetSensor 
      Caption         =   "Get"
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   2160
      Top             =   8400
   End
   Begin MSComDlg.CommonDialog ComDial 
      Left            =   4320
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   3240
      Top             =   8640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox tbEvents 
      Height          =   3495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   7680
      Width           =   5535
   End
   Begin VB.CommandButton butServer 
      Caption         =   "Enable Server"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "ID"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Sensor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Bin"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Number"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1980
      Width           =   615
   End
End
Attribute VB_Name = "frmDiscover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const mConMaxConnections As Integer = 3
Const PortNum As Long = 1500
Dim ServerEnabled As Boolean

Private Sub butGetBin_Click()
    On Error GoTo ErrHandler
    SendPacket pdGetBinNum, "", "28 59 af 73 6 0 0 b"
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butGetBin_Click", Err.Description
     Resume ErrExit

End Sub

Private Sub butGetSensor_Click()
    On Error GoTo ErrHandler
    SendPacket pdGetSenNum, "", "28 59 af 73 6 0 0 b"
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butGetSensor_Click", Err.Description
     Resume ErrExit
End Sub

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

Private Sub butSetBin_Click()
    On Error GoTo ErrHandler
    SendPacket pdSetBinNum, txtBin, "28 59 af 73 6 0 0 b"
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butSetBin_Click", Err.Description
     Resume ErrExit

End Sub

Private Sub butSetSensor_Click()
    On Error GoTo ErrHandler
    SendPacket pdSetSenNum, txtSensor, "28 59 af 73 6 0 0 b"
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butSetSensor_Click", Err.Description
     Resume ErrExit

End Sub

Private Sub butUpdate_Click()

    On Error GoTo ErrHandler
    'tell all monitors to update all sensors
    SendPacket pdReport
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butUpdate_Click", Err.Description
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

Private Sub butSignal_Click()
    On Error GoTo ErrHandler
    SendPacket pdGetSignalStrength
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmUpdate", "butSignal_Click", Err.Description
     Resume ErrExit
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

Private Sub ProcessData(ND As String)
    Dim DataType As PacketData
    Dim NewVal As String
    Dim ID As String
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
        If UBound(Parts) = 3 Then
            DataType = Val(Parts(0))
            NewVal = Parts(1)
            ID = Parts(2)
            Select Case DataType
                Case PacketData.pdTemp
                    EventMessage "Temperature = " & NewVal
                Case PacketData.pdGetBinNum
                    EventMessage "bin # = " & NewVal
                Case PacketData.pdGetSenNum
                    EventMessage "sensor # = " & NewVal
                Case PacketData.pdGetSignalStrength
                    EventMessage "RSSI = " & NewVal
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

Private Sub SendPacket(DataType As PacketData, Optional NewVal As String = "", Optional ID As _
    String = "", Optional Client As Integer = -1)
    Dim S As String
    Dim Sck As Winsock
    On Error GoTo ErrHandler
    Debug.Print NewVal
    If ServerEnabled Then
        CheckClients
        S = BeginPacket & DataType & "|" & NewVal & "|" & ID & "|"
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

Private Sub Timer1_Timer()
    SendPacket pdHeartBeat
End Sub

