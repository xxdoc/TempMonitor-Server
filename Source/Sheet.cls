VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mGrd As MSFlexGrid       ' Private reference to an MSFlexGrid control
Attribute mGrd.VB_VarHelpID = -1
Private WithEvents mTxt As TextBox          ' Private reference to a TextBox control
Attribute mTxt.VB_VarHelpID = -1

Private mIntEditRow As Integer              ' The row being edited
Private mIntEditCol As Integer              ' The column being edited
                                            
Private Const mConRowOffset As Integer = 2  ' Row and column offsets to account for
Private Const mConColOffset As Integer = 1  ' fixed rows and columns
                                            ' Default column width and alignment
Private Const mConDefaultWidth As Integer = 1200
Private Const mConDefaultAlign As Integer = flexAlignRightCenter

Public Event CellChanged(intRow As Integer, intCol As Integer)
Public Event ValidateRow(OldRow As Integer, ReturnCol As Integer, Cancel As Boolean)
Public Event ColumnChange(OldCol As Integer)

Private OldRow As Integer
Private OldCol As Integer
Private modLoading As Boolean

Private Sub BeginEdit(Optional KeyAscii As Integer)
    ' Begin editing the currently selected cell.
    Dim intOffset As Integer
    
    intOffset = 30          ' Tweakable parameter adjusting for lack of border.
    mIntEditRow = Row       ' Save the position of the cell being edited.
    mIntEditCol = Col
    
    With mTxt        ' Configure the edit TextBox
    
        ' Process the passed key to create familiar spreadsheet behavior.
        Select Case KeyAscii
        
        ' If no key passed, copy current cell contents into edit text box.
        Case 0
            .Text = mGrd.Text
        
        ' Clear cell contents.
        Case vbKeyBack, vbKeyDelete
            .Text = ""
        
        ' Otherwise, replace current text in cell.
        Case Else
            .Text = Chr$(KeyAscii)
            .SelStart = 1
            
        End Select
        
        ' Size and show the edit text box.
        .Left = mGrd.Left + mGrd.ColPos(mIntEditCol + mConColOffset) + 2 * intOffset
        .Top = mGrd.Top + mGrd.RowPos(mIntEditRow + mConRowOffset) + intOffset
        .Width = mGrd.ColWidth(mIntEditCol + mConColOffset) - (2 * intOffset)
        .Height = mGrd.RowHeight(mIntEditRow + mConRowOffset) - intOffset
        .Visible = True
        .SetFocus
        
    End With
    
End Sub

Public Sub CancelEdit()
    ' Cancel the current edit.
    mTxt.Text = ""
    mTxt.Visible = False
    
    ' Cancel edit location.
    mIntEditRow = -1
    mIntEditCol = -1
    
    ' Make sure focus is on grid
    mGrd.SetFocus
End Sub

Public Property Let Cell(intRow As Integer, intCol As Integer, vrnCellVal As Variant)
    ' Set the contents of a cell.
    ' Take a variant so that it can accept most data types.
    
    ' Check for out-of-range
    If intRow < 1 Or intRow > Rows Then Exit Property
    If intCol < 1 Or intCol > Cols Then Exit Property
    
    ' Assign value
    mGrd.TextMatrix(intRow + mConRowOffset, intCol + mConColOffset) = Format(vrnCellVal, mGrd.TextMatrix(1, intCol + mConColOffset))
    
    ' Mark row and column as changed
    RowChanged(intRow) = True
    ColChanged(intCol) = True
    
    RaiseEvent CellChanged(intRow, intCol)
End Property

Public Property Get Cell(intRow As Integer, intCol As Integer) As Variant
    ' Return the contents of a cell.
    ' Returns a variant so that the cell data can automatically be converted to
    ' receiving variable type.
    
    ' Check for out-of-range
    If intRow < 1 Or intRow > Rows Then Exit Property
    If intCol < 1 Or intCol > Cols Then Exit Property
    
    Cell = mGrd.TextMatrix(intRow + mConRowOffset, intCol + mConColOffset)
End Property

Public Sub ClearSelectedCells()
    ' Clear the selected cells
    Dim intR1 As Integer, intR2 As Integer, intC1 As Integer, intC2 As Integer
    Dim I As Integer, J As Integer
    
    ' Sort the corners of the selection region.
    With mGrd
        If .Row > .RowSel Then
            intR1 = .RowSel - mConRowOffset
            intR2 = .Row - mConRowOffset
        Else
            intR1 = .Row - mConRowOffset
            intR2 = .RowSel - mConRowOffset
        End If
        
        If .Col > .ColSel Then
            intC1 = .ColSel - mConColOffset
            intC2 = .Col - mConColOffset
        Else
            intC1 = .Col - mConColOffset
            intC2 = .ColSel - mConColOffset
        End If
        
        ' Clear all cells
        For I = intR1 To intR2
            For J = intC1 To intC2
                Cell(I, J) = ""
            Next J
        Next I
    
    End With
End Sub

Public Property Let Col(intCol As Integer)
    ' Set the active column

    ' Check for a valid column
    If intCol < 1 Or intCol > Cols Then Exit Property
    
    mGrd.Col = intCol + mConColOffset
End Property

Public Property Get Col() As Integer
    ' Get the active column
    Col = mGrd.Col - mConColOffset
End Property

Public Property Let ColAlign(intCol As Integer, intAlign As Integer)
    ' Save the column alignment
    
    ' Check for a valid column
    If intCol < 0 Or intCol > Cols Then Exit Property
    
    ' Set alignment
    mGrd.ColAlignment(intCol + mConColOffset) = intAlign
End Property

Public Property Get ColChanged(intCol As Integer) As Boolean
    ' Return contents of hidden 'column changed' row.
    
    ' Check for valid column
    If intCol < 1 Or intCol > Cols Then Exit Property

    ColChanged = CBool(mGrd.TextMatrix(0, intCol + mConColOffset))
End Property

Public Property Let ColChanged(intCol As Integer, blnChanged As Boolean)
    ' Set the contents of hidden 'column changed' row.
    
    ' Check for valid column
    If intCol < 1 Or intCol > Cols Then Exit Property
    
    mGrd.TextMatrix(0, intCol + mConColOffset) = blnChanged
End Property

Public Property Let ColFormat(intCol As Integer, strFormat As String)
    ' Set the column format string

    ' Check for a valid column
    If intCol < 1 Or intCol > Cols Then Exit Property
    
    ' Column formats are stored in row 1
    mGrd.TextMatrix(1, intCol + mConColOffset) = strFormat
End Property

Public Property Get Cols() As Integer
    ' Return the number of columns.  Actual minus number fixed.
    Cols = mGrd.Cols - mConColOffset - 1
End Property

Public Property Let Cols(intCols As Integer)
    ' Sets the number of available grid columns.
    Dim I As Integer, intCurrCols As Integer
    
    ' Set the number of columns in the sheet.  Enforce a minimium of 1.
    If intCols < 1 Then intCols = 1
    intCurrCols = Cols          ' Actual grid columns are greater by number of fixed columns
    mGrd.Cols = intCols + mConColOffset + 1
    
    ' Set default formats for any new columns.
    For I = intCurrCols + 1 To intCols
        ColChanged(I) = False
        ColAlign(I) = mConDefaultAlign
        ColWidth(I) = mConDefaultWidth
        ColTitle(I) = "Column " & CStr(I)
    Next I

End Property
Public Property Get ColTitle(intCol As Integer) As String
    ColTitle = mGrd.TextMatrix(mConRowOffset, intCol + mConColOffset)
End Property
Public Property Let ColTitle(intCol As Integer, strTitle As String)
    ' Save the column title
    
    ' Check for a valid column
    If intCol < 0 Or intCol > Cols Then Exit Property
    
    ' Set title
    mGrd.TextMatrix(mConRowOffset, intCol + mConColOffset) = strTitle
End Property

Public Property Let ColWidth(intCol As Integer, intWidth As Integer)
    ' Set the column width
    
    ' Check for a valid column
    If intCol < 0 Or intCol > Cols Then Exit Property
    
    ' Set the width
    mGrd.ColWidth(intCol + mConColOffset) = intWidth
End Property

Private Sub CommitEdit()
    ' Update the currently edited cell.
    If mTxt.Visible <> True Then Exit Sub
    
    ' Copy text back in to cell, and hide edit box.
    Cell(mIntEditRow, mIntEditCol) = mTxt.Text
    mTxt.Visible = False
    
    ' Cancel edit location.
    mIntEditRow = -1
    mIntEditCol = -1

    ' Make sure focus is on grid
    mGrd.SetFocus
End Sub

Public Sub CopyToClipboard()
    ' Copy the selected cells to the clipboard.
    Clipboard.SetText mGrd.Clip
End Sub

Public Property Set EditBox(txtEdit As TextBox)
    ' Set the private TextBox variable
    Set mTxt = txtEdit
    
    ' Set some properties
    With mTxt
        .Visible = False
        .Text = ""
        .TabStop = False
        
        ' Note - This properties must be set at design time.
        '.BorderStyle = vbBSNone
    End With
End Property

Public Sub FillDown()
    ' Copy the contents of the first row in the selection
    ' down to the subsequent rows.
    Dim I As Integer, J As Integer, intStart As Integer
    Dim strVal As String
    
    intStart = mGrd.Row
    For I = mGrd.Col To mGrd.ColSel
        For J = mGrd.Row To mGrd.RowSel
            If J = intStart Then
                ' Save first value
                strVal = Cell(J - mConRowOffset, I - mConColOffset)
            Else
                ' Set values below
                Cell(J - mConRowOffset, I - mConColOffset) = strVal
            End If
        Next J
    Next I

End Sub

Public Sub FillRight()
    ' Copy the contents of the first column in the selection
    ' across to the subsequent columns.
    Dim I As Integer, J As Integer, intStart As Integer
    Dim strVal As String
    
    intStart = mGrd.Col
    For I = mGrd.Row To mGrd.RowSel
        For J = mGrd.Col To mGrd.ColSel
            If J = intStart Then
                ' Save first value
                strVal = Cell(I - mConRowOffset, J - mConColOffset)
            Else
                ' Set values across
                Cell(I - mConRowOffset, J - mConColOffset) = strVal
            End If
        Next J
    Next I
    
End Sub

Public Function Find(Optional strString As Variant) As Boolean
    ' Search through the grid starting at the current location and look
    ' for cell contents matching the given string.
    Static strSearch As String, lngPos As Long
    Dim I As Long, lngCells As Long, intRow As Integer, intCol As Integer, intCols As Integer
    
    Find = False        ' Assume not found
    
    ' If search string is given then assume the search is starting new
    If Not IsMissing(strString) Then
        strSearch = strString   ' Store search string and initial position
        lngPos = 0&             ' in static variables
    End If
    
    lngCells = mGrd.Rows * mGrd.Cols - 1&               ' Calculate the upper looping bound
    intCols = mGrd.Cols                                 ' Save the number of columns

    ' Loop over all cells searching for the string
    For I = lngPos To lngCells - 1&                     ' Loop over all cells
        intRow = (I \ intCols) - mConRowOffset          ' Calculate row and column from
        intCol = (I Mod intCols) - mConColOffset        ' sequential index
        If intRow > 0 And intCol > 0 Then               ' Exclude fixed rows and columns
            If InStr(mGrd.TextArray(I), strSearch) Then ' Look for string
                MakeCellVisible intRow, intCol          ' Make cell visible when found
                lngPos = I + 1                          ' Store next starting position
                Find = True                             ' Return success
                Exit Function
            End If
        End If
    Next I
    
End Function
Public Property Get GridWidth() As Long
    Dim C As Long
    Dim W As Long
    For C = 1 To mGrd.Cols - 1
        W = W + mGrd.ColWidth(C)
    Next C
    GridWidth = W
End Property

Public Property Set Grid(grdFlex As MSFlexGrid)
    Dim I As Integer
    
    ' Set the private MSFlexGrid variable
    Set mGrd = grdFlex
    
    ' Set some properties
    With mGrd
        .Appearance = flexFlat
        .FocusRect = flexFocusHeavy
        .TabStop = False
        .ScrollBars = flexScrollBarBoth
        
        ' Check for proper size and format
        If .Cols < mConColOffset + 2 Then .Cols = mConRowOffset + 2
        If .Rows < mConRowOffset + 2 Then .Rows = mConRowOffset + 2
        If .FixedCols < mConColOffset + 1 Then .FixedCols = mConColOffset + 1
        If .FixedRows < mConRowOffset + 1 Then .FixedRows = mConRowOffset + 1
        
        .ColWidth(0) = 0            ' Hide column dedicated to saving changed row state.
        .ColWidth(1) = 500          ' Set first fixed column width.
        .RowHeight(0) = 0           ' Hide row dedicated to saving changed column state.
        .RowHeight(1) = 0           ' Hide row dedicated to saving the column format string.
        
        ' Initialize column formats and changed state to false
        For I = 1 To Cols
            ColChanged(I) = False
            ColAlign(I) = mConDefaultAlign
            ColWidth(I) = mConDefaultWidth
            ColTitle(I) = "Column " & CStr(I)
        Next I
                
        ' Initialize the row title and changed state to false
        For I = 1 To Rows
            RowChanged(I) = False
            RowTitle(I) = CStr(I) & "."     ' Row Title
        Next I
        
    End With
    OldRow = mConRowOffset + 1
    OldCol = mConColOffset + 1
End Property

Public Sub MakeCellVisible(intRow As Integer, intCol As Integer)
    ' Make the given cell visible by scrolling as necessary
    
    ' Check for out-of-range
    If intRow < 1 Or intRow > Rows Then Exit Sub
    If intCol < 1 Or intCol > Cols Then Exit Sub
    
    ' Select the cell
    Row = intRow
    Col = intCol
    
    ' If the row or column is not visible then make it visible
    If mGrd.RowIsVisible(intRow + mConRowOffset) = False Then mGrd.TopRow = intRow + mConRowOffset
    If mGrd.ColIsVisible(intCol + mConColOffset) = False Then mGrd.LeftCol = intCol + mConColOffset
    
    ' Make sure grid has the focus
    mGrd.SetFocus
End Sub

Private Sub mGrd_DblClick()
    ' Begin editing a cell when it is double-clicked.
    BeginEdit
End Sub

Private Sub mGrd_GotFocus()
    ' Commit an edit when the grid gets focus.
    CommitEdit
End Sub

Private Sub mGrd_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Override certain keycode processing for enhanced navigation.
    Dim blnShift As Boolean, blnCtrl As Boolean, blnAlt As Boolean
    
    ' Save the state of combination keys.
    blnShift = (Shift And vbShiftMask) > 0
    blnAlt = (Shift And vbAltMask) > 0
    blnCtrl = (Shift And vbCtrlMask) > 0
    
    With mGrd
        Select Case KeyCode
        ' Move up or down with the return key
        Case vbKeyReturn
            If blnShift = True Then
                MoveUp
            Else
                MoveDown
            End If
        ' Move left or right with the tab key
        Case vbKeyTab
            If blnShift = True Then
                MoveLeft
            Else
                MoveRight
            End If
        End Select
    End With
    
End Sub

Private Sub mGrd_KeyPress(KeyAscii As Integer)
    ' React to pressed keys.
    Select Case KeyAscii
    Case 9, 13, 27              ' Cancel Tab, CR and Esc
        KeyAscii = 0
    Case Else                   ' Pass other characters to cell editing.
        BeginEdit KeyAscii
    End Select
    
End Sub

Private Sub mGrd_LeaveCell()
    ' Commit an edit when leaving a cell.
    CommitEdit
End Sub

Private Sub mGrd_RowColChange()
    If mGrd.Row <> OldRow Then Validate
    If mGrd.Col <> OldCol Then
        RaiseEvent ColumnChange(OldCol)
        OldCol = mGrd.Col
    End If
End Sub
Public Sub Validate()
    Dim Cancel As Boolean
    Dim ReturnColumn As Integer
    If Not modLoading And RowChanged(OldRow - mConRowOffset) Then
        RaiseEvent ValidateRow(OldRow - mConRowOffset, ReturnColumn, Cancel)
        If Cancel Then
            mGrd.Row = OldRow
            If ReturnColumn > 0 Then mGrd.Col = ReturnColumn + mConColOffset
        End If
    End If
    OldRow = mGrd.Row
End Sub
Private Sub mGrd_Scroll()
    ' Commit an edit when the grid is scrolled.
    CommitEdit
End Sub

Private Sub MoveDown()
    ' Attempt to move the current cell to the one below.
    With mGrd
        
        If Row < Rows Then
            ' Edit is committed when row is changed
            Row = Row + 1
            ' Make sure the new row is visible
            If mGrd.RowIsVisible(Row + mConRowOffset) = False Then mGrd.TopRow = mGrd.TopRow + 1
        Else
            CommitEdit
        End If
        
    End With
    
End Sub

Private Sub MoveLeft()
    ' Attempt to move the current cell to the one to the left.
    With mGrd
        If Col > 1 Then
            ' Edit is committed when column is changed
            Col = Col - 1
            ' Make sure the new col is visible
            If mGrd.ColIsVisible(Col + mConColOffset) = False Then mGrd.LeftCol = mGrd.LeftCol - 1
        Else
            CommitEdit
        End If
    End With
End Sub

Private Sub MoveRight()
    ' Attempt to move the current cell to the one to the right.
    With mGrd
        If Col < Cols Then
            ' Edit is committed when column is changed
            Col = Col + 1
            ' Make sure the new col is visible
            If mGrd.ColIsVisible(Col + mConColOffset) = False Then mGrd.LeftCol = mGrd.LeftCol + 1
        Else
            CommitEdit
        End If
    End With
End Sub

Private Sub MoveUp()
    ' Attempt to move the current cell to the one below.
    With mGrd
        
        If Row > 1 Then
            ' Edit is committed when row is changed
            Row = Row - 1
            ' Make sure the new row is visible
            If mGrd.RowIsVisible(Row + mConRowOffset) = False Then mGrd.TopRow = mGrd.TopRow - 1
        Else
            CommitEdit
        End If

    End With
End Sub

Private Sub mTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    ' React to editing keys pressed while in edit text box.
    Dim blnShift As Boolean, blnCtrl As Boolean, blnAlt As Boolean
    
    ' Save the state of comination keys.
    blnShift = (Shift And vbShiftMask) > 0
    blnAlt = (Shift And vbAltMask) > 0
    blnCtrl = (Shift And vbCtrlMask) > 0
    
    ' Move the current cell or cancel edit.
    With mGrd
        Select Case KeyCode
        ' Move up or down on return or down arrow
        Case vbKeyReturn, vbKeyDown
            If blnShift = True Then
                MoveUp
            Else
                MoveDown
            End If
        ' Move up or down on up arrow
        Case vbKeyUp
            If blnShift = True Then
                MoveDown
            Else
                MoveUp
            End If
        ' Move left or right on tab
        Case vbKeyTab
            If blnShift = True Then
                MoveLeft
            Else
                MoveRight
            End If
        ' Cancel editing in escape
        Case vbKeyEscape
            CancelEdit
        End Select
    End With
    
End Sub

Private Sub mTxt_KeyPress(KeyAscii As Integer)
    ' React to pressed keys.
    Select Case KeyAscii
    Case 9, 13, 27              ' Cancel Tab, CR and Esc
        KeyAscii = 0
    End Select
End Sub

Private Sub mTxt_LostFocus()
    ' Commit any changes if the edit box loses focus.
    CommitEdit
End Sub

Public Sub PasteFromClipboard()
    ' Paste text on clipboard into sheet
    mGrd.Clip = Clipboard.GetText
End Sub

Public Property Let Row(intRow As Integer)
    ' Set the active row
    
    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    mGrd.Row = intRow + mConRowOffset
End Property

Public Property Get Row() As Integer
    ' Get the active row
    Row = mGrd.Row - mConRowOffset
End Property

Public Property Get RowChanged(intRow As Integer) As Boolean
    ' Return contents of hidden 'row changed' column.
    
    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    RowChanged = CBool(mGrd.TextMatrix(intRow + mConRowOffset, 0))
End Property

Public Property Let RowChanged(intRow As Integer, blnChanged As Boolean)
    ' Set the contents of hidden 'row changed' column.
    
    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    mGrd.TextMatrix(intRow + mConRowOffset, 0) = blnChanged
End Property

Public Property Let RowData(intRow As Integer, lngValue As Long)
    ' Set the row's long data
    
    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    ' Set data value
    mGrd.RowData(intRow + mConRowOffset) = lngValue
End Property

Public Property Get RowData(intRow As Integer) As Long
    ' Get the row's long data
    
    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    ' Get data value
    RowData = mGrd.RowData(intRow + mConRowOffset)
End Property

Public Property Let RowHeight(intRow As Integer, intHeight As Integer)
    ' Set the row height
    
    ' Check for a valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    ' Set the height
    mGrd.RowHeight(intRow + mConRowOffset) = intHeight
End Property

Public Property Get Rows() As Integer
    ' Return the number of sheet rows.  Actual minus fixed rows.
    Rows = mGrd.Rows - mConRowOffset - 1
End Property

Public Property Let Rows(intRows As Integer)
    ' Set the number of available grid rows.
    Dim I As Integer, intCurrRows As Integer
    
    ' Set the number of rows in the sheet.  Minimium of 1.
    If intRows < 1 Then intRows = 1
    intCurrRows = Rows                      ' Increase actual rows by number of fixed rows
    mGrd.Rows = intRows + mConRowOffset + 1
    
    ' Initialize rows
    For I = intCurrRows + 1 To intRows
        RowChanged(I) = False               ' Not changed
        RowTitle(I) = CStr(I) & "."         ' Row Title
    Next I
    
End Property

Public Property Let RowTitle(intRow As Integer, strTitle As String)
    ' Save the row title
    
    ' Check for a valid row
    If intRow < 1 Or intRow > Rows Then Exit Property
    
    ' Set title
    mGrd.TextMatrix(intRow + mConRowOffset, mConColOffset) = strTitle
End Property

Public Sub SelectRow(intRow As Integer)
    ' Select the given row and make it visible

    ' Check for valid row
    If intRow < 1 Or intRow > Rows Then Exit Sub

    Row = intRow                            ' Set starting selection position
    Col = 1
    
    With mGrd                               ' Set ending selection position
        .RowSel = intRow + mConRowOffset
        .ColSel = Cols + mConColOffset
        
        .TopRow = intRow + mConRowOffset    ' Make the row visible
        .SetFocus                           ' Make sure the grid has focus
    End With
    
End Sub

