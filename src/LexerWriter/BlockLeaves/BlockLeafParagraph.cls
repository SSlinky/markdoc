VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
OOption Explicit
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

' Private Backing Store
'-------------------------------------------------------------------------------
Private mMyParagraph As Paragraph
Private mContentText As String
Private mIndentLevel As Long
Private mInlineConts As List


' Properties
'-------------------------------------------------------------------------------
Public Property Let ContentText(var As String)
    If Len(mContentText) > 0 Then mContentText = mContentText & vbNewLine
    mContentText = mContentText & var
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