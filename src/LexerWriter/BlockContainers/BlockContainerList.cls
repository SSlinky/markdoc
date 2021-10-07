VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.LexerWriter.BlockContainers")
Option Explicit
Implements IBlock
Implements IBlockList
Implements IBlockContainer

'-------------------------------------------------------------------------------
'   Class: BlockContainerList
'
'   1. Changing the bullet or ordered list delimiter starts a new list.
'
'   2. A list can interrupt a paragraph. That is, no blank line is needed to
'      separate a paragraph from a following list
'
'   3. Lists can be nested within lists through indentation.
'
'   4. List item text can be separated by one or more blank lines as long as the
'      text that continues is indented to the same level.
'
'   5. To force a non-paragraph space, end the previous line with double space.
'      Indentation on the next line is not required when using this form. A code
'      block cannot be started in this format but a quote can. *shrug*
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mChildren As List
Private mIndentLevel As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Get Children() As List
    Set Children = mChildren
End Property

Public Property Let IndentationLevel(var As Long)
'   Can only be set once. Closest thing we'll get to a readonly modifier.
    If mIndentLevel = 0 Then mIndentLevel = var
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

Public Property Get IBlockList_Children() As List
    Throw = Errs.NotImplementedException
End Property

Public Property Let IBlockList_IndentationLevel(var As Long)
    IndentationLevel = var
End Property

Public Property Get IBlockList_IndentationLevel() As Long
    IBlockList_IndentationLevel = IndentationLevel
End Property

Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = Children
End Property
