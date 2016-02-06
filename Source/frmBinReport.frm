VERSION 5.00
Object = "{D9794188-4A19-4061-A8C1-BCC9E392E6DB}#3.1#0"; "dcCombo.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBinReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bin Reports"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "frmBinReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTP 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   51118083
      CurrentDate     =   42393
   End
   Begin VB.CommandButton butUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   1440
      Width           =   4080
   End
   Begin VB.CommandButton butToday 
      Caption         =   "Today"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   915
      Width           =   1200
   End
   Begin VB.CommandButton butWeek 
      Caption         =   "This Week"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   915
      Width           =   1200
   End
   Begin VB.CommandButton butMonth 
      Caption         =   "This Month"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   915
      Width           =   1200
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Range"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Single"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin dcCombo.dcComboControl Combo 
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   360
      TabIndex        =   7
      Top             =   2040
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Graph"
      TabPicture(0)   =   "frmBinReport.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dcPlot1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Table"
      TabPicture(1)   =   "frmBinReport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   5895
         Left            =   -74640
         TabIndex        =   13
         Top             =   600
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   10398
         _Version        =   393216
      End
      Begin TemperatureMonitor.dcPlot dcPlot1 
         Height          =   6135
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10821
      End
   End
   Begin dcCombo.dcComboControl Combo 
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSComCtl2.DTPicker DTP 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   15
      Top             =   1440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   51118083
      CurrentDate     =   42393
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sensor"
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bin"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "frmBinReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butMonth_Click()
    On Error GoTo ErrHandler
    DTP(0).Value = DateAdd("m", -1, Now)
    DTP(1).Value = Now
    Opt(1).Value = True
    DTP(1).Enabled = True
    butUpdate_Click
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "butMonth_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butToday_Click()
    On Error GoTo ErrHandler
    DTP(0).Value = Now
    DTP(1).Value = Now
    Opt(0).Value = True
    DTP(1).Enabled = False
    butUpdate_Click
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "butToday_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butUpdate_Click()
    On Error GoTo ErrHandler
    LoadGraph
    LoadGrid
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "butUpdate_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub butWeek_Click()
    On Error GoTo ErrHandler
    DTP(0).Value = DateAdd("ww", -1, Now)
    DTP(1).Value = Now
    Opt(1).Value = True
    DTP(1).Enabled = True
    butUpdate_Click
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "butWeek_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub Combo_Change(Index As Integer)
    If Index = 0 Then
        ComboLoadSensors (Combo(0).RecordID)
        Combo(1).MoveToRecord , 0
    End If
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
    'sensors
    Combo(1).Clear
    Set Obj = Nothing
    Set Col = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmBinReport", "ComboLoad", Err.Description
    Resume ErrExit
End Sub

Private Sub ComboLoadSensors(BinID As Long)
    Dim objSensor As clsSensor
    Dim ColSensors As clsSensors
    'sensors
    On Error GoTo ErrHandler
    Combo(1).Clear
    Set ColSensors = New clsSensors
    ColSensors.Load , BinID
    For Each objSensor In ColSensors
        Combo(1).AddItem objSensor.Number, objSensor.ID
    Next
    Combo(1).ObjectID = List_Objects.LOsensors
    Set objSensor = Nothing
    Set ColSensors = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "ComboLoadSensors", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_Load()
    AD.LoadFormData Me
    butWeek_Click
    ComboLoad
    Combo(0).MoveToRecord , 0
    SetupGrid
    butToday_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub

Private Function Hr(HourData As Long) As String
    Select Case HourData
        Case 0
            Hr = "12 AM"
        Case 1 To 11
            Hr = Format(HourData) & " AM"
        Case 12
            Hr = "12 PM"
        Case 13 To 23
            Hr = Format(HourData - 12) & " PM"
    End Select
End Function

Private Sub LoadGraph()
    Dim Count As Long
    Dim RS As Recordset
    Dim SQL As String
    Dim StDate As Date
    Dim EndDate As Date
    On Error GoTo ErrHandler
    StDate = DTP(0).Value
    EndDate = DTP(1).Value
    If Opt(0) Then
        'single date
        EndDate = DateAdd("d", 1, StDate)
    End If
    If DateDiff("d", StDate, EndDate) < 2 Then
        'daily chart
        SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "h" & Chr$(34) & ",recdate) as Xval"
        SQL = SQL & " From tblRecords"
        SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
        SQL = SQL & " and recSenID = " & Combo(1).RecordID
        SQL = SQL & " group by datepart(" & Chr$(34) & "h" & Chr$(34) & ",recdate)"
        With dcPlot1
            .Xcaption = "by Hour"
            .XlabelWidth = 700
            .XlabelFormat = "Hour"
            .YlabelFormat = "#0.0"
        End With
    ElseIf DateDiff("d", StDate, EndDate) < 8 Then
        'weekly
        SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "d" & Chr$(34) & ",recdate) as Xval"
        SQL = SQL & " From tblRecords"
        SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
        SQL = SQL & " and recSenID = " & Combo(1).RecordID
        SQL = SQL & " group by datepart(" & Chr$(34) & "d" & Chr$(34) & ",recdate)"
        With dcPlot1
            .Xcaption = "by Day"
            .XlabelWidth = 1900
            .XlabelFormat = "00"
            .YlabelFormat = "#0.0"
        End With
    Else
        'monthly
        SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "m" & Chr$(34) & ",recdate) as Xval"
        SQL = SQL & " From tblRecords"
        SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
        SQL = SQL & " and recSenID = " & Combo(1).RecordID
        SQL = SQL & " group by datepart(" & Chr$(34) & "m" & Chr$(34) & ",recdate)"
        With dcPlot1
            .Xcaption = "by Week"
            .XlabelWidth = 1900
            .XlabelFormat = "00"
            .YlabelFormat = "#0.0"
        End With
    End If
    Set RS = MainDB.OpenRecordset(SQL)
    dcPlot1.Cls
    Count = 0
    Do Until RS.EOF
        With dcPlot1
            .Record = Count
            .Xdat = RS!xval
            .Ydat = RS!Yval
        End With
        Count = Count + 1
        RS.MoveNext
    Loop
    With dcPlot1
        .Ycaption = "Temperature"
        .Title = "Bin Temperature"
        .Draw
    End With
    Set RS = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "GetData", Err.Description
     Resume ErrExit
End Sub

Private Sub LoadGrid()
    Dim RS As Recordset
    Dim SQL As String
    Dim Row As Long
    Dim Col As Long
    Dim StDate As Date
    Dim EndDate As Date
    Dim RSbins As Recordset
    Dim ID As Long
    Dim Fmat As Long
    On Error GoTo ErrHandler
    Grid.Clear
    StDate = DTP(0).Value
    EndDate = DTP(1).Value
    If Opt(0) Then
        'single date
        EndDate = DateAdd("d", 1, StDate)
    End If
    SQL = "select * from tblSensors where senStorID = " & Combo(0).RecordID
    Set RSbins = MainDB.OpenRecordset(SQL)
    Col = 0
    Do Until RSbins.EOF
        Col = Col + 1
        ID = RSbins!SenID
        If DateDiff("d", StDate, EndDate) < 2 Then
            'daily chart (by hour)
            SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "h" & Chr$(34) & ",recdate) as Xval"
            SQL = SQL & " From tblRecords"
            SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
            SQL = SQL & " and recSenID = " & ID
            SQL = SQL & " group by datepart(" & Chr$(34) & "h" & Chr$(34) & ",recdate)"
            Fmat = 0
        ElseIf DateDiff("d", StDate, EndDate) < 8 Then
            'weekly (by day)
            SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "d" & Chr$(34) & ",recdate) as Xval"
            SQL = SQL & " From tblRecords"
            SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
            SQL = SQL & " and recSenID = " & ID
            SQL = SQL & " group by datepart(" & Chr$(34) & "d" & Chr$(34) & ",recdate)"
            Fmat = 1
        Else
            'monthly (by week)
            SQL = "select avg(recTemp) as Yval, DatePart(" & Chr$(34) & "ww" & Chr$(34) & ",recdate) as Xval"
            SQL = SQL & " From tblRecords"
            SQL = SQL & " where recdate > " & ToAccessDate(StDate) & " and recdate < " & ToAccessDate(EndDate)
            SQL = SQL & " and recSenID = " & ID
            SQL = SQL & " group by datepart(" & Chr$(34) & "ww" & Chr$(34) & ",recdate)"
            Fmat = 2
        End If
        Set RS = MainDB.OpenRecordset(SQL)
        SetupGrid
        Row = 0
        Do Until RS.EOF
            Row = Row + 1
            If (Grid.Rows - 1) < Row Then Grid.Rows = Row
            Select Case Fmat
                Case 0
                    Grid.TextMatrix(Row, 0) = Hr(RS!xval)
                Case 1
                    Grid.TextMatrix(Row, 0) = RS!xval
                Case 2
                    Grid.TextMatrix(Row, 0) = "Week " & RS!xval
            End Select
            Grid.TextMatrix(Row, Col) = Format(RS!Yval, "##0.0")
            RS.MoveNext
        Loop
        RSbins.MoveNext
    Loop
    Set RSbins = Nothing
    Set RS = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmBinReport", "LoadGrid", Err.Description
     Resume ErrExit
End Sub

Private Sub Opt_Click(Index As Integer)
    If Index = 0 Then
        'single
        DTP(1).Enabled = False
    Else
        'range
        DTP(1).Enabled = True
    End If
End Sub

Private Sub SetupGrid()
    Dim C As Long
    With Grid
        .Cols = 9
        .FixedCols = 0
        .TextMatrix(0, 0) = "Date"
        .ColAlignment(0) = flexAlignCenterCenter
        .ColWidth(0) = 1000
        For C = 1 To 8
            .TextMatrix(0, C) = "Sensor " & Str(C)
            .ColAlignment(C) = flexAlignRightCenter
            .ColWidth(C) = 900
        Next C
        .Width = 8550
        .Rows = 25
    End With
End Sub

