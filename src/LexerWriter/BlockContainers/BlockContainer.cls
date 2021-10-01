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


' Private Backing Store
'-------------------------------------------------------------------------------
Private mChildren As List
Private mContentText As String
Private mIndentLevel As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Let ContentText(var As String)
    If Len(mContentText) > 0 Then mContentText = mContentText & vbNewLine
    mContentText = mContentText & var
End Property

Public Property Get ContentText() As String
    ContentText = mContentText
End Property

Public Property Get Children() As List
    Set Children = mChildren
End Property

Public Property Let IndentationLevel(var As Long)
    mIndentLevel = var
End Property

Public Property Get IndentationLevel() As Long
    IndentationLevel = mIndentLevel
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mChildren = New List
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
    Utils.WriteContent toDoc, Me
End Sub

Public Sub StyleContent(styleDoc As Document)
    Utils.StyleContent styleDoc, Me
End Sub

' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Let IBlock_IndentationLevel(var As Long)
    IndentationLevel = var
End Property

Public Property Get IBlock_IndentationLevel() As Long
    IBlock_IndentationLevel = IndentationLevel
End Property

Public Sub IBlock_WriteContent(toDoc As Document)
    WriteContent toDoc
End Sub

Public Sub IBlock_StyleContent(styleDoc As Document)
    StyleContent styleDoc
End Sub

Public Property Let IBlockContent_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockContent_ContentText() As String
    IBlockContent_ContentText = ContentText
End Property

Public Property Get IBlockContainer_Children() As List
'   Get accessor to the container's children.
    Set IBlockContainer_Children = Children
End Property
