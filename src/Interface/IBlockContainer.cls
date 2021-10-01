VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockContainer
'   A container block is a block that has other blocks as its contents. There
'   are two basic kinds of container blocks: block quotes and list items. Lists
'   are meta-containers for list items.
'
'   We define the syntax for container blocks recursively. The general form of
'   the definition is:
'
'       If X is a sequence of blocks, then the result of transforming X in
'       such-and-such a way is a container of type Y with these blocks as its
'       content.
'
'   So, we explain what counts as a block quote or list item by explaining how
'   these can be generated from their contents.
'-------------------------------------------------------------------------------


Public Property Get Children() As List
'   Get accessor to the container's children.
    Throw = Errs.InterfaceUsedAsObject
End Property
