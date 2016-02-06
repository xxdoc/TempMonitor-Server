Attribute VB_Name = "Lists"
Option Explicit
'Column Properties: caption|width|alignment
'right 1, centre 2, left other
'Searches: caption,search column,IsString,Value,Visible
'GridData: GD(row,column)=String value

Public Enum List_Objects
    LOsensors
    LOclients
    LObins
    LOmaps
End Enum

Private locFormNames() As String
Private locDataObjectNames() As String
Private locGridCaptions() As String

Public Sub BuildNameLists()
'---------------------------------------------------------------------------------------
' Procedure : BuildNameLists
' Date      : 12/Mar/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Dim S As String
    S = "frmSensors,frmClients,frmStorage,frmBinMaps"
    locFormNames = Split(S, ",")
    S = "clsSensor,clsClient,MapDisplay,StorageDisplay"
    locDataObjectNames = Split(S, ",")
    S = "Sensors,Clients,Bins,Bin Maps"
    locGridCaptions = Split(S, ",")
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "BuildNameLists", Err.Description
    Resume ErrExit
End Sub

Public Function DataObjectName(ByVal ID As Long) As String
    DataObjectName = locDataObjectNames(ID)
End Function

Public Function FormName(ByVal ID As Long) As String
    FormName = locFormNames(ID)
End Function

Public Function GetForm(Form_Name As String) As Form
'---------------------------------------------------------------------------------------
' Procedure : GetForm
' Author    : David
' Date      : 03/Feb/2010
' Purpose   :
'---------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    Select Case LCase(Form_Name)
        Case "frmsensors"
            Set GetForm = frmSensors
        Case "frmclients"
            Set GetForm = frmClients
        Case "frmstorage"
            Set GetForm = frmStorage
        Case "frmbinmaps"
            Set GetForm = frmBinMaps
    End Select
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "GetForm", Err.Description
    Resume ErrExit
End Function

Public Function GetObject(ObjectName As String) As Object
'---------------------------------------------------------------------------------------
' Procedure : GetObject
' Author    : David
' Date      : 03/Feb/2010
' Purpose   :
'---------------------------------------------------------------------------------------

    On Error GoTo ErrHandler
    Select Case LCase(ObjectName)
        Case "clsensor"
            Set GetObject = New clsSensor
        Case "clsclient"
            Set GetObject = New clsClient
        Case "storage"
            Set GetObject = New StorageDisplay
        Case "map"
            Set GetObject = New MapDisplay
    End Select
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "GetObject", Err.Description
    Resume ErrExit
End Function

Public Function ListData(ListID As List_Objects, Optional DataRequired As Long _
    = 3) As String()
'---------------------------------------------------------------------------------------
' Procedure : DisplayList
' Date      : 29/Mar/2010
' Purpose   : datarequired = 1 column properties, 2 - search properties,
' 3 - grid data, 4 - sort column
' updated Dec 26, 2014 to use enums
'---------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Select Case ListID
        Case List_Objects.LOsensors
            ListData = ListSensors(DataRequired)
        Case List_Objects.LOclients
            ListData = ListClients(DataRequired)
        Case List_Objects.LObins
            ListData = ListBins(DataRequired)
        Case List_Objects.LOmaps
            ListData = ListMaps(DataRequired)
    End Select
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "DisplayList", Err.Description
    Resume ErrExit
End Function

Public Function ListGridCaption(ByVal ID As Long) As String
    ListGridCaption = locGridCaptions(ID)
End Function

Public Function ListSensors(DataRequired As Long) As String()
'---------------------------------------------------------------------------------------
' Procedure : ListSensors
' Date      : 22/Dec/15
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim objCollection As New clsSensors
    Dim dataObj As clsSensor
    Dim I As Long
    Dim ColProps(0) As String
    Dim GridData() As String
    Dim RowCount As Long
    Dim ColumnCount As Long
    Dim SP(0) As String
    On Error GoTo ErrHandler
    Select Case DataRequired
        Case 1
           'build column properties
            ColProps(0) = "ID|0|0"
            ColProps(0) = ColProps(0) & "|Record #|1000|2"
            ColProps(0) = ColProps(0) & "|Bin |2000|2"
            ColProps(0) = ColProps(0) & "|Sensor Number|1500|2"
            ColProps(0) = ColProps(0) & "|Rom Code|2000|2"
            ColProps(0) = ColProps(0) & "|Description|2000|2"
            ColProps(0) = ColProps(0) & "|Client|2000|2"
            ColProps(0) = ColProps(0) & "|Temperature|1000|2"
            ListSensors = ColProps
        Case 2
            'searches
            SP(0) = "Bin Number,2,-1,-1,-1"
            SP(0) = SP(0) & ",Rom Code,4,-1,0,-1"
            ListSensors = SP
        Case 3
            'grid data
            ColumnCount = 8
            ReDim GridData(0, ColumnCount - 1) 'to prevent error on no data
            objCollection.Load
            'check for no data
            RowCount = objCollection.Count
            If RowCount > 0 Then
                ReDim GridData(RowCount - 1, ColumnCount - 1)
                For I = 0 To RowCount - 1
                    Set dataObj = objCollection.Item(I + 1)
                    With dataObj
                        GridData(I, 0) = .ID
                        GridData(I, 1) = .RecNum
                        GridData(I, 2) = .BinDescription
                        GridData(I, 3) = .Number
                        GridData(I, 4) = .RomCode
                        GridData(I, 5) = .Description
                        GridData(I, 6) = .ClientDescription
                        GridData(I, 7) = .MaxTemp
                    End With
                Next I
            End If
            ListSensors = GridData
        Case 4
            'sort column
            SP(0) = "0"
            ListSensors = SP
    End Select
    Set objCollection = Nothing
    Set dataObj = Nothing
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "ListSensors", Err.Description
    Resume ErrExit
End Function

Public Function ListClients(DataRequired As Long) As String()
'---------------------------------------------------------------------------------------
' Procedure : ListClients
' Date      : 04/Jan/16
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim objCollection As New clsClients
    Dim dataObj As clsClient
    Dim I As Long
    Dim ColProps(0) As String
    Dim GridData() As String
    Dim RowCount As Long
    Dim ColumnCount As Long
    Dim SP(0) As String
    On Error GoTo ErrHandler
    Select Case DataRequired
        Case 1
           'build column properties
            ColProps(0) = "ID|0|0"
            ColProps(0) = ColProps(0) & "|Record #|1000|2"
            ColProps(0) = ColProps(0) & "|Mac Address|2000|2"
            ColProps(0) = ColProps(0) & "|Description|2000|2"
            ListClients = ColProps
        Case 2
            'searches
            SP(0) = "Description,3,-1,-1,-1"
            SP(0) = SP(0) & ",Mac Address,2,-1,0,-1"
            ListClients = SP
        Case 3
            'grid data
            ColumnCount = 4
            ReDim GridData(0, ColumnCount - 1) 'to prevent error on no data
            objCollection.Load
            'check for no data
            RowCount = objCollection.Count
            If RowCount > 0 Then
                ReDim GridData(RowCount - 1, ColumnCount - 1)
                For I = 0 To RowCount - 1
                    Set dataObj = objCollection.Item(I + 1)
                    With dataObj
                        GridData(I, 0) = .ID
                        GridData(I, 1) = .RecNum
                        GridData(I, 2) = .Mac
                        GridData(I, 3) = .Description
                    End With
                Next I
            End If
            ListClients = GridData
        Case 4
            'sort column
            SP(0) = "0"
            ListClients = SP
    End Select
    Set objCollection = Nothing
    Set dataObj = Nothing
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "ListClients", Err.Description
    Resume ErrExit
End Function

Public Function ListMaps(DataRequired As Long) As String()
    Dim objCollection As New Maps
    Dim dataObj As MapDisplay
    Dim I As Long
    Dim ColProps(0) As String
    Dim GridData() As String
    Dim RowCount As Long
    Dim ColumnCount As Long
    Dim SP(0) As String
    On Error GoTo ErrHandler
    Select Case DataRequired
        Case 1
           'build column properties
            ColProps(0) = "ID|0|0"
            ColProps(0) = ColProps(0) & "|Record #|1000|2"
            ColProps(0) = ColProps(0) & "|Name|2000|2"
            ListMaps = ColProps
        Case 2
            'searches
            SP(0) = "Record #,1,0,-1,-1"
            SP(0) = SP(0) & ",Name,2,-1,0,-1"
            ListMaps = SP
        Case 3
            'grid data
            ColumnCount = 3
            ReDim GridData(0, ColumnCount - 1) 'to prevent error on no data
            objCollection.Load
            'check for no data
            RowCount = objCollection.Count
            If RowCount > 0 Then
                ReDim GridData(RowCount - 1, ColumnCount - 1)
                For I = 0 To RowCount - 1
                    Set dataObj = objCollection.Item(I + 1)
                    With dataObj
                        GridData(I, 0) = .ID
                        GridData(I, 1) = .RecNum
                        GridData(I, 2) = .MapName
                    End With
                Next I
            End If
            ListMaps = GridData
        Case 4
            'sort column
            SP(0) = "0"
            ListMaps = SP
    End Select
    Set objCollection = Nothing
    Set dataObj = Nothing
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "ListMaps", Err.Description
    Resume ErrExit
End Function


Public Function ListBins(DataRequired As Long) As String()
    Dim objCollection As New Storages
    Dim dataObj As StorageDisplay
    Dim I As Long
    Dim ColProps(0) As String
    Dim GridData() As String
    Dim RowCount As Long
    Dim ColumnCount As Long
    Dim SP(0) As String
    On Error GoTo ErrHandler
    Select Case DataRequired
        Case 1
           'build column properties
            ColProps(0) = "ID|0|0"
            ColProps(0) = ColProps(0) & "|Storage #|1000|2"
            ColProps(0) = ColProps(0) & "|Description|3000|2"
            ListBins = ColProps
        Case 2
            'searches
            SP(0) = "Storage #,1,0,-1,-1"
            SP(0) = SP(0) & ",Description,2,-1,0,-1"
            ListBins = SP
        Case 3
            'grid data
            ColumnCount = 3
            ReDim GridData(0, ColumnCount - 1) 'to prevent error on no data
            objCollection.Load
            'check for no data
            RowCount = objCollection.Count
            If RowCount > 0 Then
                ReDim GridData(RowCount - 1, ColumnCount - 1)
                For I = 0 To RowCount - 1
                    Set dataObj = objCollection.Item(I + 1)
                    With dataObj
                        GridData(I, 0) = .ID
                        GridData(I, 1) = .Label
                        GridData(I, 2) = TR(.Description)
                    End With
                Next I
            End If
            ListBins = GridData
        Case 4
            'sort column
            SP(0) = "1"
            ListBins = SP
    End Select
    Set objCollection = Nothing
    Set dataObj = Nothing
ErrExit:
    Exit Function
ErrHandler:
    AD.DisplayError Err.Number, "Lists", "ListBins", Err.Description
    Resume ErrExit
End Function



