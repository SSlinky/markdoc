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
Implements IBlockContainer
Implements IDocWritable

'-------------------------------------------------------------------------------
'   Class: BlockContainerList
'-------------------------------------------------------------------------------

Private mChildren As New List

Public Property Get IBlockContainer_Children() As List
    Set IBlockContainer_Children = mChildren
End Property

Public Property Let IBlockContainer_Content(textContent As String)
    Throw = Errs.NotImplementedexception
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

Public Sub WriteContent()
'   This method MUST be called before StyleContent.

'   Writes the document by recursively calling the WriteContent method
'   of all docWritable objects in mChildren. There is no error handling as
'   any object in mChildren that doesn't implement IDocWritable should throw.

    Dim docWritable As IDocWritable
    For Each docWritable In mChildren.Items
        docWritable.WriteContent
    Next
End Sub

Public Sub StyleContent()
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

