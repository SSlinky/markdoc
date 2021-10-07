VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Document")
Option Explicit

'-------------------------------------------------------------------------------
'   Class: InlineContent
'   Container to hold text, style, and document location information.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mText As String
Private mRange As Range
Private mStyle As Style


' Properties
'-------------------------------------------------------------------------------
Public Property Let Text(var As String)
    mText = var
End Property

Public Property Get Text() As String
    Text = mText
End Property

Public Property Set Range(var As Range)
    Set mRange = var
End Property

Public Property Get Range() As Range
    With mRange
        If .End <> .Start + Len(mText) Then
            Dim docLength As Long
            Dim rangeEnd As Long
            
            docLength = Len(.Parent.content)
            rangeEnd = .Start + Len(mText)
            
            Set mRange = .Parent.Range(.Start, Utils.Min(rangeEnd, docLength))
        End If
    End With
    Set Range = mRange
End Property

Public Property Set Style(var As Style)
    Set mStyle = var
End Property

Public Property Get Style() As Style
    Set Style = mStyle
End Property
