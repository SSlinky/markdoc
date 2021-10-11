VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockContentInline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockContentInline
'   Identifies the block as having inline content.
'-------------------------------------------------------------------------------

Public Property Set InlineContents(var As List)
Attribute InlineContents.VB_Description = "The parsed inline content."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get InlineContents() As List
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let ContentText(var As String)
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get ContentText() As String
    Throw = Errs.InterfaceUsedAsObject
End Property
