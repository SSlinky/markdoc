VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Document")
Option Explicit
Implements ISortable
Implements IInlineContent

'-------------------------------------------------------------------------------
'   Class: InlineCode
'   Container to hold text, style, and document location information.
'   Adding text to this container processes further code only rules.
'
'   The contents of the code span are normalised in the following ways:

'   * First, line endings are converted to spaces.
'   * If the resulting string both begins and ends with a space character, but
'     does not consist entirely of space characters, a single space character is
'     removed from the front and back.
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
'   Replace line endings with spaces if they exist.
'   This intentionally happens before the second check.
    var = Replace(var, vbCrLf, " ")

'   Trim the first and last char from the val if conditions met.
    Dim allSpaces As Boolean
    Dim leftRightIsSpace As Boolean

    allSpaces = var = Space(Len(var))
    leftRightIsSpace = Left(var, 1) & Right(var, 1) = "  "
    
    If leftRightIsSpace And Not allSpaces Then
        var = Mid(var, 2, Len(var) - 2)
    End If

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
    LazySetStyle
    Set Style = mStyle
End Property

Public Property Let SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort number for this content."
    mSortIndex = var
End Property

Public Property Get SortIndex() As Long
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
    Text = Text & str
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub LazySetStyle()
Attribute LazySetStyle.VB_Description = "Sets the style if it needs to."
'   Sets the style if it needs to.'
'
    If mStyle Is Nothing Then
        Set mStyle = DocumentShortcuts.GetStyle("InlineCode")
    ElseIf Not mStyle.NameLocal = "InlineCode" Then
        Set mStyle = DocumentShortcuts.GetStyle("InlineCode")
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
