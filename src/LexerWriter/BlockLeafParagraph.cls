VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafParagraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlockLeaf
Implements IDocWritable

'-------------------------------------------------------------------------------
'   Class: BlockLeafParagraph
'
'   A leaf block paragraph is a block that has content and styling information.
'   Content and style may be spread over a series of
'-------------------------------------------------------------------------------


Private Type StyledContent
    Text As String
    Range As Range
End Type


' Properties
'-------------------------------------------------------------------------------
Public Property Let IBlockLeaf_Content(textContent As String)
    Throw = Errs.NotImplementedException
End Property


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub IBlockLeaf_WriteContent()
    Throw = Errs.NotImplementedException
End Sub

Public Sub IBlockLeaf_StyleContent()
    Throw = Errs.NotImplementedException
End Sub

Public Sub IDocWritable_WriteContent()
    Throw = Errs.NotImplementedException
End Sub

Public Sub IDocWritable_StyleContent()
    Throw = Errs.NotImplementedException
End Sub


' Methods
'-------------------------------------------------------------------------------
