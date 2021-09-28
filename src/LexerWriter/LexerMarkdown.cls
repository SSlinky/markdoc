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
    Dim docLevelBlock As BlockContainer
    Set docLevelBlock = mBlockStack.Items(mBlockStack.Count)
    ParseBlockTree docLevelBlock
End Sub

Public Sub WriteDocument()
'   Writes the parsed markdown to a new document.

    Dim mkdwn As BlockContainer
    Set mkdwn = mBlockStack.Items(mBlockStack.Count)

    Dim doc As Document
    Set doc = Application.Documents.Add
    
    mkdwn.WriteContent(doc)
    mkdwn.StyleContent(doc)
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
    Dim indetLevel As Long

    CloseBlock
    indentLevel = ThisBlock.IndentationLevel

    OpenChildBlock
    ThisBlock.IndentationLevel = indentLevel
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
Private Sub ParseLineType(line As String)
'   Prepares the line and calls ParseLineTypeAction
'
'   Args:
'       line: The raw line as it is read from the source.

    Dim cleanLine As String
    Dim unindentLine As String
    Dim thisIndentLevel As Long

'   Clean, get indentation, and unindent line.
    cleanLine = Utils.CleanString(line)

    If Trim(cleanLine) = vbNullString Then
        thisIndentLevel = ThisBlock.IndentationLevel
    Else
        unindentLine = Utils.UnindentLine(cleanLine, ThisBlock.IndentationLevel)
        thisIndentLevel = Utils.GetIndentLevel(unindentLine)
    End If

    ParseLineTypeAction cleanLine, unindentLine, thisIndentLevel
End Sub

Private Sub ParseLineTypeAction( _
    lineCleaned As String, _
    lineUnindented As String, _
    indentLevel As Long)
'   Performs the parse action depending on current state.
'   This sub can call itself recursively if more than one action relevant.
'
'   Args:
'       lineCleaned: The cleaned version of the original line.
'       lineUnindented: The unindented version of the cleaned line.
'       indentLevel: The indent level in spaces of the line.
'
'   Raises:
'
    Dim thisContentBlock As IBlockContent

'   An indentation level reduction will always kick us out of the block.
'   Except when the 
    If indentLevel < ThisBlock.IndentationLevel Then
        CloseBlockLeaf
        ParseLineTypeAction lineCleaned, lineUnindented, indentLevel
    End If

    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencededCodeLeaf
'           Check for a close fence string.
            If Parser.IsCodeBlockFence(cleanLine, mFencedCodeFence) Then
                CloseBlockLeaf
'           Nothing else pulls us out of this block type.                
            Else
                Set thisContentBlock = ThisBlock
                thisFencedCodeBlock.ContentText = lineUnindented
            End If

'       Parsing from inside indented code block.
        Case Is = IsInIndentedCodeLeaf
            Set thisContentBlock = ThisBlock
            thisFencedCodeBlock.ContentText = lineUnindented

'       Parsing from inside a paragraph.
        Case Is = IsInParagraphLeaf
'           Blank lines break the paragraph.
            If unindentLine = vbNullString Then
                OpenSiblingBlock
                Utils.CBlockLeaf ThisBlock, New BlockLeafBlankLine
'           A code fence breaks the paragraph.
            ElseIf Parser.IsCodeBlockFence(unindentLine, mFencedCodeFence)
                OpenSiblingBlock
                Utils.CBlockLeaf ThisBlock, New BlockLeafFencedCode
                ThisBlock.ContentText = unindentLine
'           A new list breaks the paragraph.
            ElseIf Parser.IsListItem(unindentLine)

'           Nothing else pulls us out of this block type.
            Else
                Set thisContentBlock = ThisBlock
                thisFencedCodeBlock.ContentText = lineUnindented
            End If        

'       Parsing from inside a block of blank lines.
        Case Is = IsInBlankLineLeaf
'           Non-blank line should open an unknown sibling block for parsing.
            If Not unindentLine = vbNullString Then
                OpenSiblingBlock
                ParseLineTypeAction lineCleaned, lineUnindented, indentLevel
            End If
    End Select 
End Sub

Private Sub ParseBlockTree(block As IBlock)
'   Parses the block content if it's a leaf and calls this method
'   recursively on its children. You should only need to pass the
'   document level block to this sub.
'
'   Args:
'       block: The block with content or children to be parsed.
'
'   Raises:
'       BlockTypeMismatch if passed an object that is neither
'       IBlockContainer nor IBlockContent.

    If TypeOf block Is IBlockContainer Then
        Dim container As IBlockContainer
        Dim child As IBlock

        Set container = block
        For Each childc In container.Children
            ParseBlockTree child
        Next child

    ElseIf TypeOf block Is IBlockContent Then
        Dim content As IBlockContent
        Set content = block
        Debug.Print("Parsing... " & content.ContentText)
    
    Else
        Throw = BlockTypeMismatch
    End If

End Sub


' Helpers: Parser State
'-------------------------------------------------------------------------------
Private Function IsInFencededCodeLeaf() As Boolean
    IsInFencededCodeLeaf = Not mFencedCodeFence = ""
End Function

Private Function IsInIndentedCodeLeaf() As Boolean
'   True if the current code block is an indented code block.
    IsInIndentedCodeLeaf = TypeOf ThisBlock Is IndentedCodeBlock
End Function

Private Function IsInParagraphLeaf() As Boolean
'   Parser state flag is in paragraph.
'
'   Returns:
'       True if the current block is a BlockLeafParagraph

    IsInParagraphLeaf = TypeOf ThisBlock Is BlockLeafParagraph
End Function

Private Function IsInBlankLineLeaf() As Boolean
'   Parser state flag is in blank lines.
'
'   Returns:
'       True if the current block is a BlockLeafParagraph

    IsInBlankLineLeaf = TypeOf ThisBlock Is BlockLeafBlankLine
End Function
