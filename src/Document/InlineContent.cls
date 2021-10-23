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
Implements ISortable
Implements IInlineContent

'-------------------------------------------------------------------------------
'   Class: InlineContent
'   Container to hold text, style, and document location information.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mText As String
Private mRange As Range
Private mStyle As Style
Private mSortIndex As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Let Text(var As String)
    Throw = Errs.WarnFeatureDeprecated
    mText = var
End Property

Public Property Get Text() As String
    Throw = Errs.WarnFeatureDeprecated
    Text = mText
End Property

Public Property Set Range(var As Range)
    Throw = Errs.WarnFeatureDeprecated
    Set mRange = var
End Property

Public Property Get Range() As Range
    Throw = Errs.WarnFeatureDeprecated
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
    Throw = Errs.WarnFeatureDeprecated
    Set mStyle = var
End Property

Public Property Get Style() As Style
    Throw = Errs.WarnFeatureDeprecated
    LazySetStyle
    Set Style = mStyle
End Property

Public Property Let SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort number for this content."
    Throw = Errs.WarnFeatureDeprecated
    mSortIndex = var
End Property

Public Property Get SortIndex() As Long
    Throw = Errs.WarnFeatureDeprecated
    SortIndex = mSortIndex
End Property


' Methods
'-------------------------------------------------------------------------------
Public Sub AppendText(str As String)
Attribute AppendText.VB_Description = "Appends text to the Text property."
'   Appends text to the Text property
'
'   Args:
'       str: The string to append to the Text property.
'
    Throw = Errs.WarnFeatureDeprecated
    Text = Text & str
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub LazySetStyle()
Attribute LazySetStyle.VB_Description = "Sets the style if it needs to."
'   Sets the style if it needs to.'
'
    If mStyle Is Nothing Then
        Set mStyle = DocumentShortcuts.GetStyle("Normal")
    End If
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Let ISortable_SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort index for this item."
    SortIndex = var
End Property

Public Property Get ISortable_SortIndex() As Long
    ISortable_SortIndex = SortIndex
End Property

Public Property Let IInlineContent_Text(var As String)
    Text = var
End Property

Public Property Get IInlineContent_Text() As String
    IInlineContent_Text = Text
End Property

Public Property Set IInlineContent_Range(var As Range)
    Set Range = var
End Property

Public Property Get IInlineContent_Range() As Range
    Set IInlineContent_Range = Range
End Property

Public Property Set IInlineContent_Style(var As Style)
    Set Style = var
End Property

Public Property Get IInlineContent_Style() As Style
    Set IInlineContent_Style = Style
End Property

Public Property Let IInlineContent_SortIndex(var As Long)
    SortIndex = var
End Property

Public Property Get IInlineContent_SortIndex() As Long
    IInlineContent_SortIndex = SortIndex
End Property

Public Sub IInlineContent_AppendText(str As String)
    AppendText str
End Sub
