VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockList
'   A list is a collection of list items. Lists can be ordered or unordered.
'   All items in the list must be of the same type and indentation level.
'-------------------------------------------------------------------------------


Public Property Get Children() As List
'   Getter for the container's children.
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let IndentationLevel(var As Long)
'   Setter for the number of characters this list is indented as an
'   offset from its parent object.
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get IndentationLevel() As Long
'   Getter for the number of characters this list is indented as an
'   offset from its parent object.
    Throw = Errs.InterfaceUsedAsObject
End Property
