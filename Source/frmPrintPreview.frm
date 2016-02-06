VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Begin VB.Form frmPrintPreview 
   Caption         =   "Preview Report"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   Icon            =   "frmPrintPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   2415
      Left            =   6120
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
      _cx             =   4260
      _cy             =   4260
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   11.8371212121212
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   0
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   5355
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton butFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butPrev 
         Caption         =   "<"
         Height          =   375
         Left            =   640
         TabIndex        =   11
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butNext 
         Caption         =   ">"
         Height          =   375
         Left            =   1160
         TabIndex        =   9
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butShrink 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butExpand 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton butPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton butPDF 
         Caption         =   "Export to PDF"
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton butClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   7680
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox tbPage 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Text            =   "tbPage"
         Top             =   30
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog ComDial 
      Left            =   3120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPDF8LibCtl.VSPDF8 PDF 
      Left            =   2400
      Top             =   2880
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
   Begin VSReport8LibCtl.VSReport VSReport1 
      Left            =   3840
      Top             =   2880
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoadOk As Boolean
Private ZM As Long
Private locStartFolder As String
Private modHideReportFooter As Boolean
Private modMoveLast As Boolean

Private Sub butClose_Click()
    Unload Me
End Sub
Private Sub butExpand_Click()
    If ZM < 4 Then ZM = ZM + 1
    DoZoom
End Sub
Private Sub butFirst_Click()
    VSPrinter1.PreviewPage = 1
    ShowPage
End Sub
Private Sub butLast_Click()
    VSPrinter1.PreviewPage = VSPrinter1.PageCount
    ShowPage
End Sub
Private Sub butNext_Click()
    VSPrinter1.PreviewPage = VSPrinter1.PreviewPage + 1
    ShowPage
End Sub
Private Sub butPDF_Click()
'---------------------------------------------------------------------------------------
' Procedure : butPDF_Click
' Author    : David
' Date      : 12/3/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim DestinationName As String
    On Error GoTo ErrHandler
    ComDial.Flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    ComDial.InitDir = locStartFolder
    ComDial.Filter = "PDF|*.pdf"
    ComDial.FileName = ""
    ComDial.DefaultExt = "pdf"
    ComDial.CancelError = False
    ComDial.DialogTitle = "Save As"
    ComDial.ShowSave
    DestinationName = ComDial.FileName
    If DestinationName = "" Then
        MsgBox "Invalid File Name."
    Else
        PDF.ConvertDocument VSPrinter1, DestinationName
        MsgBox "PDF document saved to " & DestinationName
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "butPDF_Click", Err.Description
    Resume ErrExit
End Sub
Private Sub butPrev_Click()
    On Error GoTo ErrHandler
    VSPrinter1.PreviewPage = VSPrinter1.PreviewPage - 1
    ShowPage
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "butPrev_Click", Err.Description
    Resume ErrExit
End Sub
Private Sub butPrint_Click()
'---------------------------------------------------------------------------------------
' Procedure : butPrint_Click
' Author    : David
' Date      : 12/24/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    VSPrinter1.PrintDoc True
'    VSPrinter1.PrintDialog (pdPrint)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "butPrint_Click", Err.Description
    Resume ErrExit
End Sub
Private Sub butShrink_Click()
    If ZM > 0 Then ZM = ZM - 1
    DoZoom
End Sub
Private Sub DoZoom()
    Select Case ZM
        Case 0
            VSPrinter1.Zoom = 50
            VSPrinter1.ZoomMode = zmPercentage
        Case 1
            VSPrinter1.Zoom = 100
            VSPrinter1.ZoomMode = zmPercentage
        Case 2
            VSPrinter1.ZoomMode = zmPageWidth
        Case 3
            VSPrinter1.ZoomMode = zmWholePage
        Case 4
            VSPrinter1.ZoomMode = zmTwoPages
    End Select
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHandler
    If modMoveLast Then
        butLast_Click
        modMoveLast = False
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmPrintPreview", "Form_Activate", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.ActiveControl.Name = "tbPage" Then
        tbPage_Validate False
        tbPage_GotFocus
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrExit
    Dim SF As New CSystemFolders
    AD.LoadFormData Me
    locStartFolder = SF.Path(CSIDL_PERSONAL)
    LoadOk = False
    ZM = Val(AD.AppData("PreviewZoom"))
    VSReport1.OnOpen = ""
    LoadOk = True
ErrExit:
End Sub
Private Sub Form_Resize()
'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : David
' Date      : 12/24/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim FW As Long
    Dim FH As Long
    Dim NewH As Long
    Dim NewW As Long
    On Error GoTo ErrHandler
    FW = Me.Width
    FH = Me.Height
    NewW = FW - 300
    NewH = FH - 240 - 1200
    If NewW < 500 Then NewW = 500
    If NewH < 500 Then NewH = 500
    VSPrinter1.Left = 120
    VSPrinter1.Top = 120
    VSPrinter1.Width = NewW
    VSPrinter1.Height = NewH
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "Form_Resize", Err.Description
    Resume ErrExit
End Sub
Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
    AD.AppData("PreviewZoom") = ZM
End Sub
Public Property Let OnOpenData(NewVal As String)
    VSReport1.OnOpen = NewVal
End Property
Public Property Let ReportFooterHide(NewVal As Boolean)
    modHideReportFooter = NewVal
End Property
Public Sub PrintTextFile(FileName As String)
    Dim Ln As String
    On Error GoTo ErrExit
    With VSPrinter1
        .Clear
        .StartDoc
        .FontBold = True
        .FontUnderline = True
        .Paragraph = "GrainManager Log"
        .FontBold = False
        .FontUnderline = False
        .Paragraph = Format(Now, "medium date")
        .Paragraph = ""
    End With
    Open FileName For Input As #1
    Do While Not EOF(1)
        Input #1, Ln
        VSPrinter1.Paragraph = Ln
    Loop
    Close #1
    VSPrinter1.EndDoc
    DoZoom
    ShowPage
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "PrintTextFile", Err.Description
    Resume ErrExit
End Sub
Public Sub PrintPic(Pic As Picture)
    On Error GoTo ErrHandler
    With VSPrinter1
        .Clear
        .StartDoc
        .DrawPicture Pic, "0in", "0in", , , vppaStretch
        .EndDoc
    End With
    DoZoom
    ShowPage
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "PrintPic", Err.Description
    Resume ErrExit
End Sub
Public Sub Render()
'---------------------------------------------------------------------------------------
' Procedure : Render
' Author    : David
' Date      : 1/4/2011
' Purpose   :
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    If modHideReportFooter Then
        VSReport1.Sections("ReportFooter").Visible = False
    End If
    If Not VSReport1.IsBusy Then
        VSReport1.Render VSPrinter1
        DoZoom
        ShowPage
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmPrintPreview", "Render", Err.Description
    Resume ErrExit
End Sub
Public Property Let ReportName(NewVal As String)
    VSReport1.Load App.Path & "\gmanager.xml", NewVal
End Property
Public Property Set ReportRecordset(NewVal As Object)
    On Error GoTo ErrHandler
    VSReport1.DataSource.Recordset = NewVal
    On Error GoTo 0
ErrExit:
    Exit Property
ErrHandler:
    Select Case Err.Number
        Case 3021
            'no data
            MsgBox "No Records."
            Unload Me
        Case Else
            AD.DisplayError Err.Number, "frmPrintPreview", "Property Set ReportRecordset", Err.Description
        Resume ErrExit
    End Select
End Property
Private Sub ShowPage()
    tbPage = VSPrinter1.PreviewPage & "/" & VSPrinter1.PageCount
End Sub
Private Sub tbPage_GotFocus()
    tbPage.SelStart = 0
    tbPage.SelLength = Len(tbPage.Text)
End Sub
Private Sub tbPage_Validate(Cancel As Boolean)
    VSPrinter1.PreviewPage = Val(tbPage)
    ShowPage
End Sub
Private Sub VSPrinter1_AfterUserPage()
    ShowPage
End Sub
Private Sub VSReport1_OnError(ByVal Number As Long, ByVal Description As String, Handled As Boolean)
    AD.DisplayError Number, "frmPrintPreview", "", Description
End Sub
Public Sub MoveLast()
    modMoveLast = True
End Sub

