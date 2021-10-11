VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISortable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: ISortable
'   An interface for a sortable list item.
'-------------------------------------------------------------------------------

Public Property Let SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort index for this item."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get SortIndex() As Long
    Throw = Errs.InterfaceUsedAsObject
End Property
