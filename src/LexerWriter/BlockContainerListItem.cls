VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockContainer
Implements IDocWritable

'-------------------------------------------------------------------------------
'   Class: BlockContainerListItem
'
'   1. A list is an ordered list if its constituent list items begin with ordered
'      list markers.
'
'   2. Two list items are of the same type if they begin with a list marker of the
'      same type. Two list markers are of the same type if they are ordered list
'      numbers with the same delimiter (either . or )).
'
'   3. The start number of an ordered list is determined by the list number of its
'      initial list item. The numbers of subsequent list items are disregarded.
'-------------------------------------------------------------------------------

' Properties
'-------------------------------------------------------------------------------
Private mChildren As List
Public Property Get Children() As List
    Set Children = mChildren
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mChildren = New List
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent()
'   This method MUST be called before StyleContent.
    Throw = Errs.NotImplementedException
End Sub

Public Sub StyleContent()
'   This method MUST be called after WriteContent.
    Throw = Errs.NotImplementedException
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = Children
End Property

Public Sub IBlockContainer_WriteContent()
    WriteContent
End Sub

Public Sub IBlockContainer_StyleContent()
    StyleContent
End Sub

Public Sub IDocWritable_WriteContent()
    WriteContent
End Sub

Public Sub IDocWritable_StyleContent()
    StyleContent
End Sub
