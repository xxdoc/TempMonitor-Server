VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TemperatureMonitor Log"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbEvents 
      Height          =   8535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim Ln As String
    Dim Mes As String
    Dim ER As Long
    On Error GoTo ErrHandler
    AD.LoadFormData Me
    Open AD.Folders(App_Folders_Common) & "\Log.txt" For Input As #1
    Do While Not EOF(1)
        Input #1, Ln
        Ln = Ln & vbNewLine
        Mes = Mes & Ln
    Loop
    Close #1
    tbEvents = Mes
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    ER = Err.Number
    Select Case ER
        Case 53
            'no file
            Call MsgBox("No data.", vbInformation, App.Title)
            Unload Me
        Case Else
            AD.DisplayError Err.Number, "frmLog", "Load", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub

Private Sub tbEvents_GotFocus()
    tbEvents.SelStart = Len(tbEvents.Text)
End Sub
