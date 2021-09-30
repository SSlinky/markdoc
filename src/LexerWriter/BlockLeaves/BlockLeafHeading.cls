VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafHeading"
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
'   Class: BlockLeafHeading
'   A block that represents a title or heading of any level.
'   ATX Headings:
'       - Regex: ^\s{0,3}(#{1,6})\s+(.*?)\s*(?:#*\s*)$
'       - The heading content is capture group 2.
'       - Content parsed as inline text, similar to a paragraph.
'       - The heading level is derived from capture group 1.
'       - Level is count of # minus one, 0 is title, > 0 is heading.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mMyParagraph As Paragraph
Private mContentText As String
Private mIndentLevel As Long
Private mInlineConts As List
Private mHeadngLevel As Long

' Properties
'-------------------------------------------------------------------------------
Public Property Let ContentText(var As String)
    mContentText = var
End Property

Public Property Get ContentText() As String
    ContentText = mContentText
End Property

Public Property Let IndentationLevel(var As Long)
    mIndentLevel = var
End Property

Public Property Get IndentationLevel() As Long
    IndentationLevel = mIndentLevel
End Property

Public Property Let Level(var As Long)
    If Utils.Between(var, 0, 6) Then
        mHeadngLevel = var
    End If
End Property

Public Property Get Level() As Long
    HeadngLevel = mHeadngLevel
End Property


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
'   This method MUST be called before StyleContent.
    If mContentText = vbNullString Then Exit Sub

    With toDoc
'       Initialise new paragraph
        If Len(.content) > 1 Then
            If Not Utils.EndsWithNewLine(.content) Then
                .content.InsertAfter Text:=vbNewLine
            End If
            .Paragraphs.Add
        End If

'       Save the working paragraph for later styling.
        Set mMyParagraph = .Paragraphs(.Paragraphs.Count)
    End With

'   Set ranges and write content to new paragraph.
'   >>>>
'   For now, directly set the content as inline content.
'   This code should be removed as functionality developed.
    Set mInlineConts = New List
    mInlineConts.Push New InlineContent
'   <<<<

    Dim inlineCont As InlineContent
    For Each inlineCont In mInlineConts
        With inlineCont
            Set .Range = toDoc.Range(Len(toDoc.content) - 1)
            .Range.Text = .Text
        End With
    Next inlineCont
End Sub

Public Sub StyleContent(styleDoc As Document)
'   This method MUST be called after WriteContent.
    Throw = Errs.NotImplementedException
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub IBlockLeaf_WriteContent(toDoc As Document)
    WriteContent toDoc
End Sub

Public Sub IBlockLeaf_StyleContent(styleDoc As Document)
    StyleContent styleDoc
End Sub

Public Sub IDocWritable_WriteContent(toDoc As Document)
    WriteContent toDoc
End Sub

Public Sub IDocWritable_StyleContent(styleDoc As Document)
    StyleContent styleDoc
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
