Attribute VB_Name = "modHelp"
Option Explicit
Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal _
    hWndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, dwData As _
    Any) As Long

Const HH_DISPLAY_TOPIC As Long = 0
Const HH_HELP_CONTEXT As Long = &HF

Public Sub DisplayHelp(ID As Long, mForm As Form)
    On Error GoTo ErrHandler
    HtmlHelp mForm.hWnd, "ProgHelp.chm", HH_HELP_CONTEXT, ByVal ID
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "modHelp", "DisplayHelp", Err.Description
    Resume ErrExit
End Sub
Public Sub LoadHelp()
    On Error GoTo ErrHandler
    App.HelpFile = App.Path & "\ProgHelp.chm"
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "modHelp", "LoadHelp", Err.Description
    Resume ErrExit
End Sub

