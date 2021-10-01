VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockContainer

'-------------------------------------------------------------------------------
'   Class: BlockContainerQuote
'   A block quote marker, optionally preceded by up to three spaces of
'   indentation, consists of (a) the character > together with a following space
'   of indentation, or (b) a single character > not followed by a space of
'   indentation.
'
'   The following rules define block quotes:
'
'       1. Basic case. If a string of lines Ls constitute a sequence of blocks
'          Bs, then the result of prepending a block quote marker to the
'          beginning of each line in Ls is a block quote containing Bs.
'
'       2. Laziness. If a string of lines Ls constitute a block quote with
'          contents Bs, then the result of deleting the initial block quote
'          marker from one or more lines in which the next character other than
'          a space or tab after the block quote marker is paragraph continuation
'          text is a block quote with Bs as its content. Paragraph continuation
'          text is text that will be parsed as part of the content of a
'          paragraph, but does not occur at the beginning of the paragraph.
'
'       3. Consecutiveness. A document cannot contain two block quotes in a row
'          unless there is a blank line between them.
'
'   Nothing else counts as a block quote.
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mChildren As List
Private mIndentLevel As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Let IndentationLevel(var As Long)
    mIndentLevel = var
End Property

Public Property Get IndentationLevel() As Long
    IndentationLevel = mIndentLevel
End Property

Public Property Get Children() As List
    Set Children = mChildren
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

Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = Children
End Property
