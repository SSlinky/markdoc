VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockContainerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IBlock
Implements IBlockList
Implements IBlockContainer
Implements IDocWritable

'-------------------------------------------------------------------------------
'   Class: BlockContainerList
'
'   1. Changing the bullet or ordered list delimiter starts a new list.
'
'   2. A list can interrupt a paragraph. That is, no blank line is needed to
'      separate a paragraph from a following list
'
'   3. Lists can be nested within lists through indentation.
'
'   4. List item text can be separated by one or more blank lines as long as the
'      text that continues is indented to the same level.
'
'   5. To force a non-paragraph space, end the previous line with double space.
'      Indentation on the next line is not required when using this form. A code
'      block cannot be started in this format but a quote can. *shrug*
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mChildren As List
Private mIndentLevel As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Get Children() As List
    Set Children = mChildren
End Property

Public Property Let IndentLevel(var As Long)
'   Can only be set once. Closest thing we'll get to a readonly modifier.
    If mIndentLevel = 0 Then mIndentLevel = var
End Property

Public Property Get IndentLevel() As Long
    IndentLevel = mIndentLevel
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set mChildren = New List
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document)
'   This method MUST be called before StyleContent.

'   Writes the document by recursively calling the WriteContent method
'   of all docWritable objects in mChildren. There is no error handling as
'   any object in mChildren that doesn't implement IDocWritable should throw.

    Dim docWritable As IDocWritable
    For Each docWritable In mChildren.Items
        docWritable.WriteContent
    Next
End Sub

Public Sub StyleContent(styleDoc As Document)
'   This method MUST be called after WriteContent.

'   Styles the document by recursively calling the StyleContent method
'   of all docWritable objects in mChildren. There is no error handling as
'   any object in mChildren that doesn't implement IDocWritable should throw.

    Dim docWritable As IDocWritable
    Do While mChildren.Items.Count > 0
        Set docWritable = mChildren.Pop
        docWritable.StyleContent
    Loop
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = Children
End Property

Public Sub IBlockContainer_WriteContent(toDoc As Document)
    WriteContent
End Sub

Public Sub IBlockContainer_StyleContent(styleDoc As Document)
    StyleContent
End Sub

Public Sub IDocWritable_WriteContent(toDoc As Document)
    WriteContent
End Sub

Public Sub IDocWritable_StyleContent(styleDoc As Document)
    StyleContent
End Sub

Public Property Get IBlockList_Children() As List
    Throw = Errs.NotImplementedException
End Property

Public Property Let IBlockList_IndentLevel(var As Long)
    IndentLevel = var
End Property

Public Property Get IBlockList_IndentLevel() As Long
    IBlockList_IndentLevel = IndentLevel
End Property

Public Property Let IBlock_IndentationLevel(var As Long)
    IndentationLevel = var
End Property

Public Property Get IBlock_IndentationLevel() As Long
    IBlock_IndentationLevel = IndentationLevel
End Property
