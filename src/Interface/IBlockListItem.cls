VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Interface")
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockListItem
'   A list item is a child of a list. It contains a link back to its parent.
'-------------------------------------------------------------------------------


Public Property Get Parent() As List
'   Getter for the item's parent object.
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get IndentationLevel() As Long
'   Getter for the number of characters this list is indented as an
'   offset from its parent object. This property links back to the parent.
    Throw = Errs.InterfaceUsedAsObject
End Property
