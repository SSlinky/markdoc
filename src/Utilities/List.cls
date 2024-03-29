VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Utilities")
Option Explicit

'-------------------------------------------------------------------------------
'   Class: List
'   Wrapper class for a VBA.Collection
'   Adds Push and Pop methods and exposes the base collection.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mBase As Collection
Private mIsTapeStyle As Boolean


' Properties
'-------------------------------------------------------------------------------
Public Property Get IsTapeStyle() As Boolean
'   Pushing and popping is FIFO.
'    - New items are pushed after last item.
'    - First item is popped.
    IsTapeStyle = mIsTapeStyle
End Property

Public Property Get IsStandardStyle() As Boolean
'   Pushing and popping is LIFO.
'    - New items are pushed before first item.
'    - First item is popped.
    IsStandardStyle = Not mIsTapeStyle
End Property

Public Property Get Peek() As Variant
'   Same as Pop but doesn't remove the object.
    If IsObject(mBase.Item(1)) Then
        Set Peek = mBase.Item(1)
    Else
        Peek = mBase.Item(1)
    End If
End Property

Public Property Get Count() As Long
    Count = mBase.Count
End Property

Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
    Set Items = mBase
End Property

Public Property Get Enumerator() As IUnknown
Attribute Enumerator.VB_Description = "Gets an enumerator that iterates through the List."
Attribute Enumerator.VB_UserMemId = -4
'Gets an enumerator that iterates through the List.
    Set Enumerator = mBase.[_NewEnum]
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mBase = New Collection
    mIsTapeStyle = True
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub Push(obj As Variant)
'   Adds an item to the list before first or after last, depending on style.
    If mIsTapeStyle Or mBase.Count = 0 Then
        mBase.Add obj
    Else
        mBase.Add obj, Before:=1
    End If
End Sub

Public Sub PushSort(obj As ISortable)
Attribute Identifier.VB_Description = "Pushes an ISortable into a sorted stack position."
'   Pushes an ISortable into a sorted stack position
'
'   Args:
'       obj: The object to be push sorted.
'
    Dim searchLow As Long
    Dim searchMid As Long
    Dim searchHigh As Long

    Dim indexLow As Long
    Dim indexMid As Long
    Dim indexObj As Long
    Dim indexHigh As Long

    Dim lastRun As Boolean

    searchLow = 1
    searchHigh = mBase.Count

'   Don't search if we have no objects.
    If searchHigh = 0 Then
        Push obj
        Exit Sub
    End If

'   Perform a binary search on the data.
    indexObj = obj.SortIndex
    Do
'       Set the middle point.
        searchMid = Int((searchHigh - searchLow) / 2) + searchLow

'       Set the indexes to test against.
        indexLow = mBase.Item(searchLow).SortIndex
        indexMid = mBase.Item(searchMid).SortIndex
        indexHigh = mBase.Item(searchHigh).SortIndex

'       Check if lower than low or higher than high.
        If indexObj <= indexLow Then
            searchMid = searchLow
            Exit Do
        ElseIf indexObj >= indexHigh Then
            searchMid = searchHigh + 1
            Exit Do
        End If

'       Exit if last run
        If lastRun Then Exit Do

'       Update the low or high boundary.
        If indexObj > indexMid Then
            searchLow = searchMid + 1
        Else
            searchHigh = searchMid - 1
        End If

        If searchHigh < searchLow Then
            Debug.Assert False
        End If

'       Flag last run or not.
        lastRun = searchLow >= searchHigh
    Loop

'   Insert at sorted position.
    Insert searchMid, obj
End Sub


Public Function Pop() As Variant
'   Returns the first item, popping it from the list.
    If IsObject(mBase.Item(1)) Then
        Set Pop = mBase.Item(1)
    Else
        Pop = mBase.Item(1)
    End If
    mBase.Remove 1
End Function

Public Sub SetTapeStyle()
'   Sets pushing and popping to FILO
    If Not mIsTapeStyle Then
        ReverseCollection
        mIsTapeStyle = True
    End If
End Sub

Public Sub SetStandardStyle()
'   Sets pushing and popping to FIFO
    If mIsTapeStyle Then
        ReverseCollection
        mIsTapeStyle = False
    End If
End Sub

Public Function PopFind(obj As Variant) As Long
Attribute IndexOf.VB_Description = "Pops an object at the location obj is found."
'   Pops an object at the location obj is found.
'
'   Args:
'       obj: The object to find.
'
'   Returns:
'       The index location of the object.
'
    Dim i As Long
    i = IndexOf(obj)
    If i > 0 Then mBase.Remove i
    PopFind = i
End Function

Public Function IndexOf(obj As Variant) As Long
Attribute IndexOf.VB_Description = "Returns the index of an object if found."
'   Returns the index of an object if found.
'
'   Args:
'       obj: The object to find.
'
'   Returns:
'       The index location of the object.
'
'   Raises:
'
    Dim i As Long
    Dim listItem As Variant

    For Each listItem In mBase
        i = i + 1
        If listItem Is obj Then
            IndexOf = i
            Exit Function
        End If
    Next listItem
    Logger.Log "Item not found in List", Level.Warning
End Function

Public Sub Insert(index As Long, obj As Variant)
'   Inserts an item at the specified index.
'
'   Args:
'       index: The index location to insert to.
'       obj: The object to be inserted.
'
    If (mBase.Count = 0 And index = 1) Or (index > mBase.Count) Then
        mBase.Add obj
    Else
        mBase.Add obj, Before:=index
    End If
End Sub

Public Sub FromArray(arr As Variant)
Attribute FromArray.VB_Description = "Load list items from an array."
'   Load list items from a single dimension array.
'
'   Args:
'       arr: A 1D array of anything.
'
    Dim lb As Long
    Dim ub As Long

    lb = LBound(arr)
    ub = UBound(arr)

    Dim i As Long
    For i = lb To ub
        Push arr(i)
    Next i
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub ReverseCollection()
    Dim cnt As Long
    cnt = mBase.Count - 1
    If cnt <= 0 Then Exit Sub

    Dim tBase As New Collection
    tBase.Add mBase.Item(1)
    mBase.Remove 1

    Dim i As Long
    For i = 1 To cnt
        tBase.Add mBase.Item(1), Before:=1
        mBase.Remove 1
    Next i

    Set mBase = tBase
End Sub
