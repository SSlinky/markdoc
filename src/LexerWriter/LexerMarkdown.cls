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
Private mIndentLevel As Long
Private mPreviousLine As String
Private mLine As String


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
    Do While Not textStream.EoF
        mLine = textStream.ReadNextLine
        ParseLineType mLine
        mPreviousLine = line
    Loop
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
        Utils.CBlockLeaf block, new BlockLeafParagraph
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
    thisIndentLevel = Utils.GetIndentLevel(unindentLine)

'   Handle indentation level changes (blank lines don't update indent).
    If TypeOf ThisBlock Is IIndentBlock And  Not Trim(cleanLine) = "" Then
        If Not thisIndentLevel = mIndentLevel Then
            HandleIndentationChange thisIndentLevel
        End If
    End If

'   Get an unindented version of the line to parse.
    unindentLine = Utils.UnindentLine(cleanLine, mIndentLevel)

    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencededLeaf
            If Parser.IsCodeBlockFence(cleanLine, mFencedCodeFence) Then
                CloseBlockLeaf
            End If

'       Parsing from inside indented code block.
        Case Is = IsInIndentedLeaf

'       Parsing from inside list block.
        Case Is = IsInListBlock
'           Will need to account for indentation level changes.
'           Probably need to move this outside the Select Case since things
'           can be both in a list and in a paragraph / code block / whatever.
'           Getting kicked out of a list level will also end any leaf block, including code.

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
'   Would prefer this was broken into small pure functions in Parser.
'   I'm thinking this whole function should be binned. Should all happen from inside case statement.
'
'   Yep. Hundred per cent. Bin it.
'
'   Args:
'       newIndentLevel: The new indentation level.
'
'   Raises:
'       ParserIndentLevelOutsideRange thrown when newIndentLevel < 0.
'
'   Fenced code blocks uber alles.
    If IsInFencededLeaf Then Exit Sub

'   Throw if indentation level less than 0.
    If thisIndentLevel < 0 Then Throw = Errs.ParserIndentLevelOutsideRange
    If newIndentLevel < 0 Then Throw = Errs.ParserIndentLevelOutsideRange

'   LISTS
'   - Only start a new list level if the previous line is blank.
'   - Text at any indent level is considered a continuation of the previous
'     list item (same line) if the previous line is not blank.
'   CODE
'   - New code blocks only start when the previous line is blank.
'   - Text less than four indents from previous line marks a new paragraph
'     unless it is blank, in which case it is a continuation of the block.

'   Right indent the block.
    If newIndentLevel > mIndentLevel Then
'       If it's less than three, do nothing.
        If newIndentLevel <= mIndentLevel + 3 Then Exit Sub

'       Test for new code block
        If Not IsInIndentedLeaf And mPreviousLine = vbNullString Then
            OpenSiblingBlock
            Utils.CBlockLeaf ThisBlock, new BlockLeafParagraph
        End If

'       Four indents is a new code block if not already in one.
    End If

'   Left indent the blocks.
    Do While mIndentLevel > newIndentLevel
'       Blank lines do not interrupt an existing indented code block.    
        If IsInIndentedLeaf And mLine = vbNullString Then Exit Do
        exit do
    Loop
End Sub

' Helpers: Parser State
'-------------------------------------------------------------------------------
Private Function IsInFencededLeaf() As Boolean
    IsInFencededLeaf = Not mFencedCodeFence = ""
End Function

Private Function IsInIndentedLeaf() As Boolean
'   True if the current code block is an indented code block.
    IsInIndentedLeaf = TypeOf ThisBlock Is IndentedCodeBlock
End Function

Private Function IsInListBlock() As Boolean
End Function
