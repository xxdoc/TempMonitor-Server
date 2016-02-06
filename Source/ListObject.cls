VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'example useage:
'   Dim LotList as ListObject
'   Dim LP as Long  'list position of item
'   Set LotList = New ListObject
'Add:
'   LotList.Add NewID,LP

Private locID() As Long
Private locListCount As Long
Private IsRemoved() As Boolean
Public Function Add(ID As Long, Optional ListPosition As Long) As Boolean
Attribute Add.VB_Description = "If ID is not on list it is added at ListPosition and returns true. If ID is on the list ListPosition returns its position in the list and Add returns false."
'---------------------------------------------------------------------------------------
' Procedure : Add
' Author    : David
' Date      : 3/7/2011
' Purpose   : returns true if a new item was added to the list
'---------------------------------------------------------------------------------------
'
    Dim L As Long
    Dim Found As Boolean
    For L = 1 To locListCount
        If Not IsRemoved(L) Then
            If ID = locID(L) Then
                Found = True
                Add = False
                ListPosition = L
                Exit For
            End If
        End If
    Next L
    If Not Found Then
        locListCount = locListCount + 1
        ReDim Preserve locID(locListCount)
        ReDim Preserve IsRemoved(locListCount)
        locID(locListCount) = ID
        Add = True
        ListPosition = locListCount
    End If
End Function

Public Property Get Count(Optional ExcludeRemoved As Boolean) As Long
Attribute Count.VB_Description = "Returns the number of items in the list."
'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : David
' Date      : 11/27/2011
' Purpose   : returns the # of items in the list including removed items.
'             If ExcludeRemoved is true then it returns the number of items
'             in the list that have not been removed.
'---------------------------------------------------------------------------------------
'
    Dim X As Long
    Dim C As Long
    If ExcludeRemoved Then
        For X = 1 To locListCount
            If Not IsRemoved(X) Then C = C + 1
        Next X
        Count = C
    Else
        Count = locListCount
    End If
End Property

Public Function FindID(ListPosition As Long, Optional ID As Long) As Boolean
Attribute FindID.VB_Description = "Returns True if ID is found. ListPosition is its position in the list."
'---------------------------------------------------------------------------------------
' Procedure : Find
' Author    : David
' Date      : 3/20/2011
' Purpose   : finds the ID at ListPosition. Returns false if not found.
'---------------------------------------------------------------------------------------
'
    If IsRemoved(ListPosition) Then
        FindID = False
        ID = 0
    Else
        If ListPosition > locListCount Then
            FindID = False
        Else
            FindID = True
            ID = locID(ListPosition)
        End If
    End If
End Function

Public Function FindPosition(ID As Long, Optional ListPosition As Long) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : Find
' Author    : David
' Date      : 3/7/2011
' Purpose   : finds the position in the list of ID. Returns false if not found.
'---------------------------------------------------------------------------------------
'
    Dim X As Long
    FindPosition = False
    For X = 1 To locListCount
        If Not IsRemoved(X) Then
            If locID(X) = ID Then
                FindPosition = True
                ListPosition = X
                Exit For
            End If
        End If
    Next
End Function

Public Function ID(ListPosition As Long) As Long
Attribute ID.VB_Description = "Returns the ID at ListPosition."
    If Not IsRemoved(ListPosition) Then
        If ListPosition <= locListCount And ListPosition > 0 Then
            ID = locID(ListPosition)
        Else
            ID = 0
        End If
    End If
End Function

Public Function Remove(Optional ID As Long, Optional ListPosition As Long) As Boolean
Attribute Remove.VB_Description = "Marks an ID as removed from the list. The positions of other ID's are not changed. "
'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : David
' Date      : 11/27/2011
' Purpose   : remove item from list. Return true if it removed.
'             Return false if it is not on the list or has already been removed.
'             Either Id or ListPosition must be specified, but not both.
'---------------------------------------------------------------------------------------
'
    Dim X As Long
    If ID > 0 Then
        For X = 1 To locListCount
            If Not IsRemoved(X) Then
                If locID(X) = ID Then
                    IsRemoved(X) = True
                    Remove = True
                    Exit For
                End If
            End If
        Next X
    Else
        If ListPosition > 0 And ListPosition <= locListCount Then
            If Not IsRemoved(ListPosition) Then
                IsRemoved(ListPosition) = True
                Remove = True
            End If
        End If
    End If
End Function
