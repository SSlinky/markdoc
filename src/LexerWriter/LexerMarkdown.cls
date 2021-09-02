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

' Private Helpers

Private Function CloseBlock() As IBlockContainer
'   Return the next ancestor from the stack
    Set CloseBlock = mBlockStack.Pop
End Function

Private Function OpenChildBlock(currentBlock As IBlockContainer) As BlockContainer
'   Create a new block container as a child of the current block
    Dim nb As New BlockContainer
    currentBlock.IBlockContainer_Children.Push nb
    
'   Push the current block onto the stack and return a new block
    mBlockStack.Push currentBlock
    Set OpenChildBlock = New BlockContainer
End Function

Private Function OpenSiblingBlock() As BlockContainer
'   Return a new block container as a child of the current parent.
    Set OpenSiblingBlock = CloseBlock.OpenChildBlock
End Function


' Constructor
Private Sub Class_Initialize()
    Set mBlockStack = New List
    mBlockStack.SetStandardStyle
    
'   This is the document level block
    mBlockStack.Push New BlockContainer
End Sub
