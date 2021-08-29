VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerListUnordered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlockContainer
Implements IDocWritable

'-------------------------------------------------------------------------------
'   Class: BlockContainerListUnordered
'
'   1. A list is an ordered list if its constituent list items begin with
'      unordered list markers.
'
'   2. Two list items are of the same type if they begin with a list marker of
'      the same type. Two list markers are of the same type if they are bullet
'      list markers using the same character (-, +, or *).
'
'   3. A list is loose if any of its constituent list items are separated by
'      blank lines, or if any of its constituent list items directly contain two
'      block-level elements with a blank line between them. Otherwise a list is
'      tight. The difference in document output is that loose lists are styled
'      with wider line spacing.
'-------------------------------------------------------------------------------
