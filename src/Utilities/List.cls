VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'   Sets pushing and popping to FILO
    IsTapeStyle = mIsTapeStyle
End Property

Public Property Get IsStandardStyle() As Boolean
'   Sets pushing and popping to FIFO
    IsStandardStyle = Not mIsTapeStyle
End Property

Public Property Get Peek() As Variant
'   Same as Pop but doesn't remove the object.
    If IsObject(mBase.Item(1)) Then
        Set Pop = mBase.Item(1)
    Else
        Pop = mBase.Item(1)
    End If
End Property

Public Property Get Count() As Long
    Count = mBase.Count
End Property

Public Property Get Items() As Collection
    Set Items = mBase
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
    If mIsTapeStyle Or mBase.Count = 0 Then
        mBase.Add obj
    Else
        mBase.Add obj, Before:=1
    End If
End Sub

Public Function Pop() As Variant
    If IsObject(mBase.Item(1)) Then
        Set Pop = mBase.Item(1)
    Else
        Pop = mBase.Item(1)
    End If
    mBase.Remove 1
End Function

Public Sub SetTapeStyle()
    If Not mIsTapeStyle Then
        ReverseCollection
        mIsTapeStyle = True
    End If
End Sub

Public Sub SetStandardStyle()
    If mIsTapeStyle Then
        ReverseCollection
        mIsTapeStyle = False
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
