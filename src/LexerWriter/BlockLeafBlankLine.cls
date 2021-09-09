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
Implements IDocWritable


' Properties
'-------------------------------------------------------------------------------
Public Property Let IBlockLeaf_Content(textContent As String)
    Throw = Errs.NotImplementedexception
End Property


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


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent()
    Throw = Errs.NotImplementedexception
End Sub

Public Sub StyleContent()
Debug.Print "styling BlockLeafBlankLine" '    Throw = Errs.NotImplementedexception
End Sub
