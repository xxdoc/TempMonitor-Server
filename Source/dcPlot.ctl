VERSION 5.00
Begin VB.UserControl dcPlot 
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   ScaleHeight     =   8565
   ScaleWidth      =   8670
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3240
      Left            =   1200
      ScaleHeight     =   3210
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   16
      Top             =   105
      Width           =   7695
   End
   Begin VB.Label lbYcaption 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lbXcaption 
      Alignment       =   2  'Center
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
      Left            =   1200
      TabIndex        =   14
      Top             =   5760
      Width           =   6615
   End
   Begin VB.Label lbYaxis 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   600
      TabIndex        =   13
      Top             =   500
      Width           =   600
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Xaxis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "dcPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type udtPlot
    Title As String * 30
    Ycaption As String * 30
    Xcaption As String * 30
    PicBox As PictureBox
    Xdat() As Currency
    Ydat() As Currency
    XlabelFormat As String
    YlabelFormat As String
End Type

Private mudtProps As udtPlot
Private flgChanged As Boolean
Private Rec As Long 'record pointer starting with 0
Const StrLen As Long = 30
Const MaxNum As Long = 1000000
Private MaxX As Currency
Private MaxY As Currency
Private MinX As Currency
Private MinY As Currency
Private XaxisCount As Long
Private XaxisWidth

Private Sub AddNewRecord()
    With mudtProps
        Rec = UBound(.Xdat) + 1
        ReDim Preserve .Xdat(Rec)
        ReDim Preserve .Ydat(Rec)
    End With
    flgChanged = True
End Sub

Public Sub Cls()
    Rec = 0
    ReDim mudtProps.Xdat(0)
    ReDim mudtProps.Ydat(0)
    mudtProps.PicBox.Cls
End Sub

Public Sub Draw()
    Dim R As Long
    mudtProps.PicBox.Cls
    GetMaxs
    If RecordCount > 1 Then
        With mudtProps.PicBox
            .ScaleHeight = MaxY - MinY
            .ScaleWidth = MaxX - MinX
            .ScaleLeft = MinX
            .ScaleTop = MaxY * -1
            .DrawWidth = 1
        End With
        mudtProps.PicBox.PSet (mudtProps.Xdat(0), -mudtProps.Ydat(0))
        For R = 1 To RecordCount - 1
            mudtProps.PicBox.Line -(mudtProps.Xdat(R), -mudtProps.Ydat(R))
        Next R
        SetYaxis
        SetXaxis
    End If
End Sub

Private Function FormatLabel(V As Variant, Fmat As String) As String
    Select Case LCase(Fmat)
        Case "hour"
            FormatLabel = Hr(CLng(V))
        Case Else
            FormatLabel = Format(V, Fmat)
    End Select
End Function

Private Sub GetMaxs()
    Dim R As Long
    MinX = MaxNum
    MinY = MaxNum
    MaxX = -MaxNum
    MaxY = -MaxNum
    For R = 0 To RecordCount - 1
        If mudtProps.Xdat(R) > MaxX Then MaxX = mudtProps.Xdat(R)
        If mudtProps.Xdat(R) < MinX Then MinX = mudtProps.Xdat(R)
        If mudtProps.Ydat(R) > MaxY Then MaxY = mudtProps.Ydat(R)
        If mudtProps.Ydat(R) < MinY Then MinY = mudtProps.Ydat(R)
    Next R
    'ensure 0 is on y scale
    If 0 > MaxY Then MaxY = 0
    If 0 < MinY Then MinY = 0
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

Public Property Get MaxXval() As Currency
    MaxXval = MaxX
End Property

Public Property Get MaxYval() As Currency
    MaxYval = MaxY
End Property

Public Property Get MinXval() As Currency
    MinXval = MinX
End Property

Public Property Get MinYval() As Currency
    MinYval = MinY
End Property

Public Property Get Record() As Long
    Record = Rec
End Property

Public Property Let Record(NewVal As Long)
    If NewVal >= 0 And Rec <> NewVal Then
        If NewVal <= UBound(mudtProps.Xdat) Then
            Rec = NewVal
        Else
            AddNewRecord
        End If
    End If
End Property

Public Property Get RecordCount() As Long
    RecordCount = UBound(mudtProps.Xdat) + 1
End Property

Private Sub SetXaxis()
    Dim TickValue As Currency
    Dim I As Long
    SizeXaxis
    TickValue = Picture1.ScaleWidth / (XaxisCount - 1)
    For I = 0 To XaxisCount - 1
        Xaxis(I) = FormatLabel(Picture1.ScaleLeft + I * TickValue, mudtProps.XlabelFormat)
    Next I
End Sub

Private Sub SetYaxis()
    Dim Cap As String
    Dim I As Long
    Dim V As String
    Dim LineCount As Long
    Dim TickValue As Currency
    LineCount = Picture1.Height / 250
    TickValue = Picture1.ScaleHeight / LineCount
    For I = 0 To LineCount - 1
        V = FormatLabel(-1 * (Picture1.ScaleTop + I * TickValue * 2), mudtProps.YlabelFormat) & " --"
        Cap = Cap & V & vbCrLf & vbCrLf
    Next I
    lbYaxis = Cap
End Sub

Private Sub SetYCaption(NewCaption As String)
    Dim Cap As String
    Dim I As Long
    Dim Ltrs() As String
    ReDim Ltrs(Len(NewCaption))
    For I = 1 To Len(NewCaption)
        Ltrs(I) = Mid$(NewCaption, I, 1) & vbCrLf
    Next I
    Cap = Join(Ltrs)
    lbYcaption = Cap
End Sub

Private Sub SizeXaxis()
    Dim I As Long
    For I = 0 To 11
         Xaxis(I).Visible = False
         Xaxis(I).AutoSize = False
    Next I
    XaxisCount = (Picture1.Width / XaxisWidth)
    If XaxisCount > 12 Then XaxisCount = 12
    For I = 0 To XaxisCount - 1
        With Xaxis(I)
            .Font.Size = 10
            .Font.Bold = False
            .Visible = True
            .Width = XaxisWidth
            .Left = Picture1.Left - XaxisWidth / 2 + (Picture1.Width / (XaxisCount - 1)) * I
            .Top = Picture1.Top + Picture1.Height + 100
            .Alignment = vbCenter
            If I = XaxisCount - 1 Then
                .Alignment = vbRightJustify
                .Left = Picture1.Left + Picture1.Width - XaxisWidth
                .AutoSize = True
                If .Width > XaxisWidth Then
                    .Width = XaxisWidth
                End If
            End If
        End With
    Next I
End Sub

Public Property Get Title() As String
    Title = Trim(mudtProps.Title)
End Property

Public Property Let Title(NewVal As String)
    If Len(NewVal) > StrLen Then NewVal = Left$(NewVal, StrLen)
    If Trim(mudtProps.Title) <> NewVal Then
        mudtProps.Title = NewVal
        flgChanged = True
        lbTitle.Caption = NewVal
    End If
End Property

Private Sub UserControl_Initialize()
    Set mudtProps.PicBox = Picture1
    Cls
    XaxisWidth = 800
    mudtProps.YlabelFormat = "###0.0"
    mudtProps.XlabelFormat = "hour"
End Sub

Private Sub UserControl_Resize()
    Dim H As Long
    Dim W As Long
    W = UserControl.Width - 200
    H = UserControl.Height
    lbTitle.Width = W
    Picture1.Width = W - 1300
    Picture1.Height = H - 1500
    lbYaxis.Height = Picture1.Height
    lbYcaption.Height = Picture1.Height
    lbXcaption.Width = Picture1.Width
    lbXcaption.Top = Picture1.Top + Picture1.Height + 475
End Sub

Public Property Get Xcaption() As String
    Xcaption = Trim(mudtProps.Xcaption)
End Property

Public Property Let Xcaption(NewVal As String)
    If Len(NewVal) > StrLen Then NewVal = Left$(NewVal, StrLen)
    If Trim(mudtProps.Xcaption) <> NewVal Then
        mudtProps.Xcaption = NewVal
        flgChanged = True
        lbXcaption.Caption = Trim(mudtProps.Xcaption)
    End If
End Property

Public Property Get Xdat() As Currency
    Xdat = mudtProps.Xdat(Rec)
End Property

Public Property Let Xdat(NewVal As Currency)
    If NewVal > MaxNum Then NewVal = MaxNum
    If NewVal < -MaxNum Then NewVal = -MaxNum
    If mudtProps.Xdat(Rec) <> NewVal Then
        mudtProps.Xdat(Rec) = NewVal
        flgChanged = True
    End If
End Property

Public Property Get XlabelFormat() As String
    XlabelFormat = Trim(mudtProps.XlabelFormat)
End Property

Public Property Let XlabelFormat(NewVal As String)
    If Len(NewVal) > StrLen Then NewVal = Left$(NewVal, StrLen)
    If Trim(mudtProps.XlabelFormat) <> NewVal Then
        mudtProps.XlabelFormat = NewVal
        flgChanged = True
    End If
End Property

Public Property Get XlabelWidth() As Long
    XlabelWidth = XaxisWidth
End Property

Public Property Let XlabelWidth(NewVal As Long)
    If NewVal > 400 And NewVal < 2100 And NewVal <> XaxisWidth Then
        XaxisWidth = NewVal
        UserControl_Resize
        Draw
    End If
End Property

Public Property Get Ycaption() As String
    Ycaption = Trim(mudtProps.Ycaption)
End Property

Public Property Let Ycaption(NewVal As String)
    If Len(NewVal) > StrLen Then NewVal = Left$(NewVal, StrLen)
    If Trim(mudtProps.Ycaption) <> NewVal Then
        mudtProps.Ycaption = NewVal
        flgChanged = True
        SetYCaption Trim(mudtProps.Ycaption)
    End If
End Property

Public Property Get Ydat() As Currency
    Ydat = mudtProps.Ydat(Rec)
End Property

Public Property Let Ydat(NewVal As Currency)
    If NewVal > MaxNum Then NewVal = MaxNum
    If NewVal < -MaxNum Then NewVal = -MaxNum
    If mudtProps.Ydat(Rec) <> NewVal Then
        mudtProps.Ydat(Rec) = NewVal
        flgChanged = True
    End If
End Property

Public Property Get YlabelFormat() As String
    YlabelFormat = Trim(mudtProps.YlabelFormat)
End Property

Public Property Let YlabelFormat(NewVal As String)
    If Len(NewVal) > StrLen Then NewVal = Left$(NewVal, StrLen)
    If Trim(mudtProps.YlabelFormat) <> NewVal Then
        mudtProps.YlabelFormat = NewVal
        flgChanged = True
    End If
End Property

