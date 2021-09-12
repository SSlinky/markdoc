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


' Public Methods
'-------------------------------------------------------------------------------
Public Sub ParseMarkdown(textStream As IIo)
'   Phase 1: Parse the text stream into blocks.


End Sub

Public Sub ParseTree()
'   Phase 2: Parse the leaf blocks.

End Sub


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mBlockStack = New List
    mBlockStack.SetStandardStyle

'   Push the document level block onto the stack.
    mBlockStack.Push New BlockContainer
    OpenChildBlock
End Sub


' Parser Helpers
'-------------------------------------------------------------------------------
Function ParseLineType(line As String) As String
    Dim cleanLine As String
    cleanLine = Utils.CleanString(line)

    Select Case True
'       Parsing from inside fenced code block.
        Case Is = IsInFencedCodeBlock
            Select Case True
                Case Is = IsFencedCodeBlockEnd(line)
                    CloseBlockLeaf
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
    
'   Must include a ~ or a `
    If InStr(line, "~") + InStr(line, "`") = 0 Then Exit Function
    
'   Simple detect code block fence. This will need to be improved later to
'   account for indentation level.
    Dim i As Long
    Dim p As Long           ' Position first detected
    Dim c As String * 1     ' Current character
    Dim f As String * 1     ' Fence character
    
    
    For i = 1 To Len(line)
        c = Mid(line, i, 1)
        Select Case c
            Case Is = " "
'               Check the fence character
                If Not f = Chr(0) Then Exit For
            Case Is = "~", "`"
                Select Case f
                    Case Is = Chr(0)
                        f = c
                        p = i
                    Case Is = c
                    Case Else
                        Exit For
                End Select
            Case Else
                Exit For
        End Select
    Next i
    
    If p = 0 Then Exit Function
    If i - p < 3 Then Exit Function
    
    mFencedCodeFence = Mid(line, p, i - p)
    mFencedCodeIndent = p - 1
    IsFencedCodeBlockStart = True
End Function

Private Function IsFencedCodeBlockEnd(line As String) As Boolean
'   Cannot be the end of a fenced code block if we're already in one.
    If Not IsInFencedCodeBlock Then Exit Function

End Function


' Parser State
'-------------------------------------------------------------------------------
Private Function IsInFencedCodeBlock() As Boolean
    IsInFencedCodeBlock = Not mFencedCodeFence = ""
End Function

Private Function IsInIndentedCodeBlock() As Boolean
End Function

Private Function IsInListBlock() As Boolean
End Function


' Private Helpers
'-------------------------------------------------------------------------------
Private Function CloseBlock() As IBlockContainer
'   Return the next ancestor from the stack.
    mCurrentBlockContent = ""
    Set CloseBlock = mBlockStack.Pop
End Function

Private Function OpenChildBlock() As BlockContainer
'   Create a new block container as a child of the current block.
    Dim newChildBlock As New BlockContainer
    Dim currentBlock As IContainerBlock

'   This will raise an error if we do not have a container block
    Set currentBlock = ThisBlock

'   Associate with current, push to stack
    currentBlock.Children.Push newChildBlock
    mBlockStack.Push newChildBlock
    Set OpenChildBlock = newChildBlock
End Function

Private Function OpenSiblingBlock() As BlockContainer
'   Return a new block container as a child of the current parent.
    Set OpenSiblingBlock = CloseBlock.OpenChildBlock
End Function

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

    tb.StyledContent.Text = mCurrentBlockContent
    CloseBlock
End Sub
