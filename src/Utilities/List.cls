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

Private mBase As Collection

Private Sub Class_Initialize()
    Set mBase = New Collection
End Sub

Private Sub Class_Terminate()
    Set mBase = Nothing
End Sub

Public Sub Push(obj As Variant)
    mBase.Add obj
End Sub

Public Property Get Pop() As Variant
    If IsObject(mBase.Item(1)) Then
        Set Pop = mBase.Item(1)
    Else
        Pop = mBase.Item(1)
    End If
    mBase.Remove 1
End Property

Public Property Get Count() As Long
    Count = mBase.Count
End Property

Public Property Get Items() As Collection
    Set Items = mBase
End Property
