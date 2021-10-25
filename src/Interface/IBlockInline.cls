VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlockInline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Interface")
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlockInline
'   Describes the base inline block.
'-------------------------------------------------------------------------------
Public Property Get Closed() As Boolean
Attribute Closed.VB_Description = "Whether this block is closed or not."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let Markdown(var As String)
Attribute Markdown.VB_Description = "The string content that was parsed."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Markdown() As String
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let PositionOpened(var As Long)
Attribute PositionOpened.VB_Description = "The position this block was opened."
'   This should be semi read-only in that can only be modified from 0
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get PositionOpened() As Long
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let PositionClosed(var As Long)
Attribute PositionClosed.VB_Description = "The position this block was closed."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get PositionClosed() As Long
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Priority() As Long
Attribute Priority.VB_Description = "The priority level for this block."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Set Range(var As Range)
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Range() As Range
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort index for this item."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get SortIndex() As Long
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Valid() As Boolean
Attribute Valid.VB_Description = "Whether this block is valid or not."
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Sub Close(truncatePosition As Long)
Attribute Close.VB_Description = "Triggers a forced block closure and truncates markdown."
    Throw = Errs.InterfaceUsedAsObject    
End Sub

Public Sub ParseCharacter(char As String)
Attribute ParseCharacter.VB_Description = "Parses a character."
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub RegisterSelf(register As InlineBlockRegister)
Attribute RegisterSelf.VB_Description = "Injects the block register so that the block can interact directly."
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub StyleContent()
Attribute StyleContent.VB_Description = "Styles content in the document."
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub WriteContent()
Attribute WriteContent.VB_Description = "Writes content to the document."
    Throw = Errs.InterfaceUsedAsObject
End Sub
