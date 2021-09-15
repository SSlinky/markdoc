VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockContent
'   Blocks that can hold text content should implement this interface.
'-------------------------------------------------------------------------------


Public Property Let ContentText(var As String)
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get ContentText() As String
    Throw = Errs.InterfaceUsedAsObject
End Property
