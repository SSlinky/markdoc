VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockLeaf
Implements IDocWritable
Implements IBlockContent

'-------------------------------------------------------------------------------
'   Class: BlockLeafParagraph
'
'   A leaf block paragraph is a block that has content and styling information.
'   Content and style may be spread over a series of
'-------------------------------------------------------------------------------


Private Type StyledContent
    Text As String
    Range As Range
End Type


' Private Backing Store
'-------------------------------------------------------------------------------
Private mMyParagraph As Paragraph
Private mContentText As String
Private mIndentLevel As Long
Private mInlineTexts As List


' Properties
'-------------------------------------------------------------------------------
Public Property Let ContentText(var As String)
    If Len(mContentText) > 0 Then mContentText = mContentText & vbNewLine
    mContentText = mContentText & var
End Property

Public Property Get ContentText() As String
    ContentText = mContentText
End Property

Public Property Let IndentLevel(var As Long)
    mIndentLevel = var
End Property

Public Property Get IndentLevel() As Long
    IndentLevel = mIndentLevel
End Property


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
'   This method MUST be called before StyleContent.
    If mContentText = vbNullString Then Exit Sub

    With toDoc
'       Initialise new paragraph
        If Len(.Content) > 1 Then
            If Not Utils.EndsWithNewLine(.Content) Then
                .Content.InsertAfter Text:=vbNewLine
            End If
            .Paragraphs.Add
        End If

'       Save the working paragraph for later styling.
        Set mMyParagraph .Paragraphs(.Paragraphs.Count)

'       Add content to the new paragraph.
'       For now, just set directly and leave this place!
        .Range(Len(.Content) - 1).Text = mContentText
        Exit Sub
        
        Dim inlineText As StyledContent
        For Each inlineText In mInlineTexts
            With inlineText
                Set .Range = .Range(Len(.Content) - 1)
                .Range.Text = .Text
            End With
        Next inlineText
    End With
End Sub

Public Sub StyleContent(styleDoc As Document)
'   This method MUST be called after WriteContent.
    Throw = Errs.NotImplementedException
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub IBlockLeaf_WriteContent(toDoc As Document)
    WriteContent
End Sub

Public Sub IBlockLeaf_StyleContent(styleDoc As Document)
    StyleContent
End Sub

Public Sub IDocWritable_WriteContent(toDoc As Document)
    WriteContent
End Sub

Public Sub IDocWritable_StyleContent(styleDoc As Document)
    StyleContent
End Sub

Public Property Let IBlockLeaf_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockLeaf_ContentText() As String
    IBlockLeaf_ContentText = ContentText
End Property

Public Property Let IBlockContent_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockContent_ContentText() As String
    IBlockContent_ContentText = ContentText
End Property

Public Property Let IBlock_IndentationLevel(var As Long)
    IndentationLevel = var
End Property

Public Property Get IBlock_IndentationLevel() As Long
    IBlock_IndentationLevel = IndentationLevel
End Property
