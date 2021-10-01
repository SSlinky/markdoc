VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockLeafBlankLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockLeaf
Implements IBlockContent

'-------------------------------------------------------------------------------
'   Class: BlockLeafBlankLine
'   This block type exists to collect blank lines. It does not have any visible
'   effect on the output other than to reduce multiple blank lines into a single
'   break.
'-------------------------------------------------------------------------------

' Properties
'-------------------------------------------------------------------------------
Public Property Let ContentText(var As String)
'   Do nothing.
End Property

Public Property Get ContentText() As String
'   Do nothing.
End Property

Public Property Let IndentationLevel(var As Long)
'   Do nothing.
End Property

Public Property Get IndentationLevel() As Long
'   Do nothing.
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
'   Do nothing.
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
'   Do nothing.
End Sub

Public Sub StyleContent(styleDoc As Document)
'   Do nothing.
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

Public Property Let IBlockLeaf_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockLeaf_ContentText() As String
    IBlockLeaf_ContentText = ContentText
End Property

Public Sub IBlockLeaf_WriteContent(toDoc As Document)
    WriteContent toDoc
End Sub

Public Sub IBlockLeaf_StyleContent(styleDoc As Document)
    StyleContent styleDoc
End Sub

Public Property Let IBlockContent_ContentText(var As String)
    ContentText = var
End Property

Public Property Get IBlockContent_ContentText() As String
    IBlockContent_ContentText = ContentText
End Property
