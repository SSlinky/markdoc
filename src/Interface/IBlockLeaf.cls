VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockLeaf
'   A leaf block is a block that has content and styling information and,
'   depending on the type, may include a collection of inline blocks.
'-------------------------------------------------------------------------------


Public Property Let ContentText(textContent As String)
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get ContentText() As String
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Sub WriteContent(toDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub StyleContent(styleDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub
