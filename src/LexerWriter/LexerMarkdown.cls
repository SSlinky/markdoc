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
Private mLinkReference As Object

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
'   Initialise the block stack.
    Set mBlockStack = New List
    mBlockStack.SetStandardStyle

'   Initialise the links reference dictionary.
    Set mLinkReference = CreateObject("Scripting.Dictionary")

'   Push the document level block onto the stack.
    mBlockStack.Push New BlockContainer
    OpenChildBlock
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub ParseMarkdown(textStream As IIo)
'   Phase 1: Parse the text stream into blocks.
    Do While Not textStream.EOF
        mLine = textStream.ReadNextLine
        ParseLineType mLine
        mPreviousLine = mLine
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
    
    mkdwn.WriteContent doc
    mkdwn.StyleContent doc
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub CloseBlock()
'   Pops the current block and discards the reference
    Dim discard As IBlock
    Set discard = mBlockStack.Pop

    Logger.Log "Closed block (" & mBlockStack.Count & ")"
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

    Logger.Log "Opened child block (" & mBlockStack.Count & ")"
End Sub

Private Sub OpenSiblingBlock()
'   Close the current block and open a child of the same parent.
    Dim indlev As Long

    CloseBlock
    indlev = ThisBlock.IndentationLevel

    OpenChildBlock
    ThisBlock.IndentationLevel = indlev

    Logger.Log "Opened sibling block (" & mBlockStack.Count & ")"
End Sub

Private Function ThisBlock() As IBlock
'   Returns the current block container.
    Dim tb As Variant
    Set tb = mBlockStack.Peek
    If TypeOf tb Is IBlock Then
        Set ThisBlock = tb
    End If
End Function

Private Function CThisBlockLeaf(castToBlock As IBlockLeaf)
'   Casts ThisBlock to another block type.
'   This helper is a shortcut to casting ThisBlock as it cannot be done
'   directly. The reference is lost so the block must be popped and pushed.
'
'   Args:
'       castToBlock: The target block to cast to.
'
    Dim block As IBlock
    Set block = mBlockStack.Pop
    Utils.CBlockLeaf block, castToBlock
    mBlockStack.Push block
End Function

Private Function ThisContentBlock() As IBlockContent
'   Gets the current block container as an IBlockContent
'   which enables use of the interface methods.
'   Returns:
'       The current block.

    Set ThisContentBlock = ThisBlock
    
End Function

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
        Utils.CBlockLeaf block, New BlockLeafParagraph
    Else
        Throw = Errs.BlockTypeMismatch
    End If
End Sub

Private Function BlockTreeHasContent(currentBlock As IBlock) As Boolean
'   Checks whether anything exists to write.
'
'   Args:
'       currentBlock: The current block to test.
'
'   Returns:
'       True if something exists to write to a document.

    Dim contentBlock As IBlockContent
    Dim containerBlock As IBlockContainer
    Dim childBlock As IBlock

'   Test the current block.
    If TypeOf currentBlock Is IBlockContent Then
        Set contentBlock = currentBlock
        If Not contentBlock.ContentText = vbNullString Then
            BlockTreeHasContent = True
            Exit Function
        End If
    End If

'   Test children.
    If TypeOf currentBlock Is IBlockContainer Then
        Set containerBlock = currentBlock
        For Each childBlock In containerBlock.Children
            If BlockTreeHasContent(childBlock) Then
                BlockTreeHasContent = True
                Exit Function
            End If
        Next childBlock
    End If
End Function

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
        unindentLine = Utils.unindentLine(cleanLine, ThisBlock.IndentationLevel)
        thisIndentLevel = Utils.GetIndentationLevel(unindentLine)
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

'   An indentation level reduction will always kick us out of the block.
'   Except when the
    If indentLevel < ThisBlock.IndentationLevel Then
        CloseBlock
        ParseLineTypeAction lineCleaned, lineUnindented, indentLevel
    End If
    
    Dim matches As Object
    Dim block As IBlock
    Dim hdg As BlockLeafHeading

    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencededCodeLeaf
'           Check for a close fence string.
            If Parser.IsCodeBlockFence(lineCleaned, mFencedCodeFence) Then
                CloseBlock
'           Nothing else pulls us out of this block type.
            Else
                ThisContentBlock.ContentText = lineUnindented
            End If

'       Parsing from inside indented code block.
        Case Is = IsInIndentedCodeLeaf
            ThisContentBlock.ContentText = lineUnindented

'       Parsing from inside a paragraph.
        Case Is = IsInParagraphLeaf
            Select Case True
'               Blank lines break the paragraph.
                Case Is = lineUnindented = vbNullString
                    OpenSiblingBlock
                    CThisBlockLeaf New BlockLeafBlankLine
'               A code fence breaks the paragraph.
                Case Is = Parser.IsCodeBlockFence( _
                    lineUnindented, mFencedCodeFence)
                    OpenSiblingBlock
                    CThisBlockLeaf New BlockLeafFencedCode
                    ThisContentBlock.ContentText = lineUnindented
'               A heading breaks the paragrap.
                Case Is = Parser.IsHeading(lineUnindented, matches)
                    CloseBlock
                    Set hdg = New BlockLeafHeading

                    hdg.ContentText = matches(0).Submatches(1)
                    hdg.HeadingLevel = Parser.ParseHeadingLevel( _
                        matches(0).Submatches(1))

                    mBlockStack.Push hdg
'               A new list breaks the paragraph. NOT YET IMPLEMENTED
                ' ElseIf Parser.IsListItem(lineUnindented) Then

'               Nothing else pulls us out of this block type.
                Case Else
                    ThisContentBlock.ContentText = lineUnindented
            End Select

'       Parsing from inside a block of blank lines.
        Case Is = IsInBlankLineLeaf
'           Non-blank line should open an unknown sibling block for parsing.
            If Not lineUnindented = vbNullString Then
                OpenSiblingBlock
                ParseLineTypeAction lineCleaned, lineUnindented, indentLevel
            End If
        
'       Parse from inside an unknown block type.
        Case Else
            Dim newBlockType As BlockType
            newBlockType = Parser.ParseUnknownBlockLine(lineUnindented, matches)
            
            Select Case newBlockType
                Case Is = BlockType.Heading
                    CThisBlockLeaf New BlockLeafHeading
                    Set hdg = ThisBlock
                    hdg.ContentText = matches(0).Submatches(1)
                    hdg.HeadingLevel = Parser.ParseHeadingLevel( _
                        matches(0).Submatches(1))
            End Select
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
        For Each child In container.Children
            ParseBlockTree child
        Next child

    ElseIf TypeOf block Is IBlockContent Then
        Dim content As IBlockContent
        Set content = block
        Logger.Log "Parsing... " & content.ContentText
    
    Else
        Throw = Errs.BlockTypeMismatch
    End If

End Sub


' Helpers: Parser State
'-------------------------------------------------------------------------------
Private Function IsInFencededCodeLeaf() As Boolean
'   Parser state flag is in fenced code.
'
'   Returns:
'       True if the current block is a BlockLeafParagraph
    IsInFencededCodeLeaf = TypeOf ThisBlock Is BlockLeafFencedCode
End Function

Private Function IsInIndentedCodeLeaf() As Boolean
'   Parser state flag is in indented code.
'
'   Returns:
'       True if the current code block is an indented code block.
    IsInIndentedCodeLeaf = TypeOf ThisBlock Is BlockLeafIndentedCode
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
