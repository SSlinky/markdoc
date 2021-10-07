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
'   Pops an object at the location obj is found.
'
'   Args:
'       obj: The object to find.
'
'   Returns:
'       The index location of the object.
'
    Dim i As Long
    Dim listItem As Variant

    For Each listItem In mBase
        i = i + 1
        If listItem Is obj Then
            mBase.Remove i
            PopFind = i
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
