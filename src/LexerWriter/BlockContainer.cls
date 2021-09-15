VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockContent
Implements IBlockContainer

'-------------------------------------------------------------------------------
'   Class: BlockContainer
'   Basic implementation of a Block.
'
'   A container block is a block that has other blocks as its contents. There
'   are two basic kinds of container blocks: block quotes and list items. Lists
'   are meta-containers for list items.
'
'   We define the syntax for container blocks recursively. The general form of
'   the definition is:
'
'       If X is a sequence of blocks, then the result of transforming X in
'       such-and-such a way is a container of type Y with these blocks as its
'       content.
'
'   So, we explain what counts as a block quote or list item by explaining how
'   these can be generated from their contents.
'-------------------------------------------------------------------------------


Private mChildren As New List
Private mContentText As String

' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Get IBlockContainer_Children() As List
'   Get accessor to the container's children.
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Sub IBlockContainer_WriteContent()
'   This class should be converted to a specific container type.
    Throw = Errs.NotImplementedException
End Sub

Public Sub IBlockContainer_StyleContent()
'   This class should be converted to a specific container type.
    Throw = Errs.NotImplementedException
End Sub

Public Property Let IBlockContent_ContentText(var As String)
    If Len(mContentText) > 0 Then mContentText = mContentText & vbNewLine
    mContentText = mContentText & var
End Property

Public Property Get IBlockContent_ContentText() As String
    ContentText = mContentText
End Property
