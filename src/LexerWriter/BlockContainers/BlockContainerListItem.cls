VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.LexerWriter.BlockContainers")
Option Explicit
Implements IBlock
Implements IBlockListItem
Implements IBlockContainer

'-------------------------------------------------------------------------------
'   Class: BlockContainerListItem
'
'   1. A list is an ordered list if its constituent list items begin with ordered
'      list markers.
'
'   2. Two list items are of the same type if they begin with a list marker of the
'      same type. Two list markers are of the same type if they are ordered list
'      numbers with the same delimiter (either . or )).
'
'   3. The start number of an ordered list is determined by the list number of its
'      initial list item. The numbers of subsequent list items are disregarded.
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mChildren As List
Private mParent As IBlockList
Private mIsCreated As Boolean
Private mContentText As String


' Properties
'-------------------------------------------------------------------------------
Public Property Get Children() As List
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
    Set Children = mChildren
End Property

Public Property Get Parent() As IBlockList
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
    Set Parent = mParent
End Property

Public Property Get IsCreated() As Boolean
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
    IsCreated = mIsCreated
End Property

Public Property Let IndentationLevel(var As Long)
    ' Do nothing or throw...?
End Property

Public Property Get IndentationLevel() As Long
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
    IndentationLevel = Parent.IndentationLevel
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
End Sub

Public Sub Create(parentBlockList As IBlockList)
    Set mChildren = New List
    Set mParent = parentBlockList
    mIsCreated = True
End Sub

' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
    Utils.WriteContent toDoc, Me
End Sub

Public Sub StyleContent(styleDoc As Document)
    If Not mIsCreated Then Throw = Errs.BlockNotCreated
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

Public Property Get IBlockListItem_Parent() As List
    Set IBlockListItem_Parent = Parent
End Property

Public Property Get IBlockListItem_IndentationLevel() As Long
    IBlockListItem_IndentationLevel = IndentationLevel
End Property

Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = Children
End Property
