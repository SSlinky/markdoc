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
private mIndentationLevel as Long   


'Properties
Public Property Get BlockStack() As List
    Set BlockStack = mBlockStack
End Property


' Public Methods
Public Sub ParseMarkdown(textStream As IIo)
'   Phase 1: Parse the text stream into blocks.


End Sub

Public Sub ParseTree()
'   Phase 2: Parse the leaf blocks.

End Sub

Function ParseLineType(line As String) As String
    Select Case True
        Case Is = IsFencedCodeBlockStart
    End Select
End Function


' Constructor
Private Sub Class_Initialize()
    Set mBlockStack = New List
    mBlockStack.SetStandardStyle

'   This is the document level block.
    Dim doc As IBlockContainer
    Set doc = New BlockContainer

'   This is the first child.
    Dim par As IBlockLeaf
    Set par = New BlockLeafParagraph

'   Push the first items onto the stack.    
    doc.Children.Push par
    mBlockStack.Push doc
    mBlockStack.Push par
End Sub


' Parser Helpers
Function ParseLineType(line As String) As String
    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencedCodeBlock
            Select Case True
                Case Is = IsFencedCodeBlockEnd(line)
                    Call CloseFencedCodeBlock
                Case Else
            End Select
'       Parsing from inside indented code block.
        Case Is = IsInIndentedCodeBlock
'       Parsing from inside list block.
        Case Is = IsInListBlock
'       Parsing from inside a paragraph.
        Case Else
            Select Case True
                Case Is = IsFencedCodeBlockStart(line)
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

Private Function IsFencedCodeBlockStart(line As String) As Boolean
'   Cannot be the start of a fenced code block if we're already in one.
    If IsInFencedCodeBlock Then Exit Function
    
End Function


' Private Functions
Private Function IsInFencedCodeBlock() As Boolean
    IsInFencedCodeBlock = Not mFencedCodeFence = ""
End Function

Private Function IsInIndentedCodeBlock() As Boolean
End Function

Private Function IsInListBlock() As Boolean
End Function

' Private Helpers
Private Function CloseBlock() As IBlockContainer
'   Return the next ancestor from the stack.
    mCurrentBlockContent = ""
    Set CloseBlock = mBlockStack.Pop
End Function

Private Function OpenChildBlock(currentBlock As IBlockContainer) As BlockContainer
'   Create a new block container as a child of the current block.
    Dim nb As IBlockContainer
    Set nb = New BlockContainer
    currentBlock.Children.Push nb
    
'   Push the current block onto the stack and return a new block.
    mBlockStack.Push currentBlock
    Set OpenChildBlock = New BlockContainer
End Function

Private Function OpenSiblingBlock() As BlockContainer
'   Return a new block container as a child of the current parent.
    Set OpenSiblingBlock = CloseBlock.OpenChildBlock
End Function

Private Function ThisBlock() As BlockContainer
'   Returns the current block container.
    Dim tb As Variant
    Set tb = mBlockStack.Peek
    If TypeOf tb Is BlockContainer Then
        Set ThisBlock = mBlockStack.Peek
    End If
End Function

Private Sub CloseFencedCodeBlock()
    Dim tb As BlockLeafFencedCode
    If Not TypeOf ThisBlock Is BlockLeafFencedCode Then
        Throw = BlockTypeMismatch
    Else
        Set tb = ThisBlock
    End If

    tb.StyledContent.Text = mCurrentBlockContent
    Call CloseBlock
End Sub
