VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LexerMarkdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: LexerMarkdown
'   Parses and tokenises markdown.
'-------------------------------------------------------------------------------


Private mBlockStack As List
Private mCurrentBlockContent As String

Private mFencedCodeFence As String
Private mFencedCodeIndent As Long
Private mIndentationLevel As Long


'Properties
'-------------------------------------------------------------------------------
Public Property Get BlockStack() As List
    Set BlockStack = mBlockStack
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mBlockStack = New List
    mBlockStack.SetStandardStyle

'   Push the document level block onto the stack.
    mBlockStack.Push New BlockContainer
    OpenChildBlock
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub ParseMarkdown(textStream As IIo)
'   Phase 1: Parse the text stream into blocks.


End Sub

Public Sub ParseTree()
'   Phase 2: Parse the leaf blocks.

End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub CloseBlock()
'   Pops the current block and discards the reference
    Dim discard As IBlock
    Set discard = mBlockStack.Pop
End Sub

Private Sub OpenChildBlock()
'   Create a new block container as a child of the current block.
    Dim newChildBlock As New BlockContainer
    Dim currentBlock As IBlockContainer

'   This will raise an error if we do not have a container block
    Set currentBlock = ThisBlock

'   Associate with current, push to stack
    currentBlock.Children.Push newChildBlock
    mBlockStack.Push newChildBlock
End Sub

Private Sub OpenSiblingBlock()
'   Close the current block and open a child of the same parent.
    CloseBlock
    OpenChildBlock
End Sub

Private Function ThisBlock() As IBlock
'   Returns the current block container.
    Dim tb As Variant
    Set tb = mBlockStack.Peek
    If TypeOf tb Is IBlock Then
        Set ThisBlock = mBlockStack.Peek
    End If
End Function

Private Sub CloseBlockLeaf()
'   Closes the current block if it is a leaf.
    Dim tb As IBlockLeaf
    If Not TypeOf ThisBlock Is IBlockLeaf Then
        Throw = BlockTypeMismatch
    Else
        Set tb = ThisBlock
    End If

'   Currently throws - need to update interface
    tb.StyledContent.Text = mCurrentBlockContent
    CloseBlock
End Sub

Private Sub SetThisBlockContent()
'   Sets the content for the current block.
'
'   If it is not already a leaf of some sort, it is cast to a paragraph.
'   This is based on the assumption that all other kinds will have been
'   detected and cast by this point.

    Dim block As IBlock
    Set block = ThisBlock
    
    If TypeOf block Is IBlockContainer Then
'       Cast to paragraph.
        Utils.CBlockLeafParagraph block
    Else
        Throw = Errs.BlockTypeMismatch
    End If
End Sub

' Helpers: Parser
'-------------------------------------------------------------------------------
Private Function ParseLineType(line As String) As String
    Dim cleanLine As String
    Dim unindentLine As String
    Dim thisIndentLevel As Long

'   Clean line and test its indentation level.
    cleanLine = Utils.CleanString(line)
    thisIndentLevel = Utils.GetIndentationLevel(unindentLine)

'   Handle indentation level changes.
    If Not thisIndentLevel = mIndentationLevel Then

    End If

'   Get an unindented version of the line to parse.
    unindentLine = Utils.UnindentLine(cleanLine)

    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencedCodeBlock
            If Parser.IsCodeBlockFence(cleanLine, mFencedCodeFence) Then
                CloseBlockLeaf
            End If

'       Parsing from inside indented code block.
        Case Is = IsInIndentedCodeBlock

'       Parsing from inside list block.
        Case Is = IsInListBlock
'           Will need to account for indentation level changes.

'       Parsing from inside a paragraph.
        Case Else
            Select Case True
                Case Is = Parser.IsCodeBlockFence(cleanLine, mFencedCodeFence)
'                   Set content of current.
'                   Open a sibling.
                Case Else
                    ParseAddToCurrent line
            End Select
    End Select
End Function

Private Sub ParseAddToCurrent(line As String)
    If mCurrentBlockContent = "" Then
        mCurrentBlockContent = line
    Else
        mCurrentBlockContent = mCurrentBlockContent & vbNewLine & line
    End If
End Sub

Private Sub HandleIndentationChange(newIndentLevel As Long)
'   Handles opening and closing of blocks based on an indentation level change.
'
'   Args:
'       newIndentLevel: The new indentation level.
'
'   Raises:
'       ParserIndentLevelOutsideRange thrown when newIndentationLevel < 0.
'
'   Throw if indentation level less than 0.
    If newIndentLevel < 0 Then
        Throw = Errs.ParserIndentLevelOutsideRange
    End If
End Sub

' Helpers: Parser State
'-------------------------------------------------------------------------------
Private Function IsInFencedCodeBlock() As Boolean
    IsInFencedCodeBlock = Not mFencedCodeFence = ""
End Function

Private Function IsInIndentedCodeBlock() As Boolean
End Function

Private Function IsInListBlock() As Boolean
End Function
