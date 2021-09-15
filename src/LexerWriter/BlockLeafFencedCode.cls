VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafFencedCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockLeaf
Implements IDocWritable
Implements IBlockContent

'-------------------------------------------------------------------------------
'   Class: BlockLeafFencedCode
'
'   A leaf block paragraph is a block that has content and styling information.
'   Fenced code blocks do not contain children or siblings.
'-------------------------------------------------------------------------------


Private Type StyledContent
    Text As String
    Range As Range
End Type


' Properties
'-------------------------------------------------------------------------------
Private mContentText As String
Public Property Let ContentText(var As String)
    If Len(mContentText) > 0 Then mContentText = mContentText & vbNewLine
    mContentText = mContentText & var
End Property

Public Property Get ContentText() As String
    ContentText = mContentText
End Property


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent()
'   This method MUST be called before StyleContent.
    Throw = Errs.NotImplementedException
End Sub

Public Sub StyleContent()
'   This method MUST be called after WriteContent.
    Throw = Errs.NotImplementedException
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub IBlockLeaf_WriteContent()
    WriteContent
End Sub

Public Sub IBlockLeaf_StyleContent()
    StyleContent
End Sub

Public Sub IDocWritable_WriteContent()
    WriteContent
End Sub

Public Sub IDocWritable_StyleContent()
    StyleContent
End Sub

Public Property Let IBlockLeaf_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockLeaf_ContentText() As String
    IBlockLeaf_ContentText = ContentText
End Property

Public Property Let IBlockContent_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockContent_ContentText() As String
    IBlockContent_ContentText = ContentText
End Property
