VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockInline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ISortable
Implements IBlockInline

'-------------------------------------------------------------------------------
'   Class: BlockInline
'   The base inline block implementation.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mClosed As Boolean
Private mMarkdown As String
Private mPosition As Long
Private mRange As Range
Private mSortIndex As Long

Private mRegister As InlineBlockRegister


' Properties
'-------------------------------------------------------------------------------
Public Property Get Closed() As Boolean
Attribute Closed.VB_Description = "Whether the block is closed or not."
    Closed = mClosed
End Property

Public Property Let Markdown(var As String)
Attribute Markdown.VB_Description = "The string content that was parsed."
'   This property appends to existing rather than overwriting.
    mMarkdown = mMarkdown & var
End Property

Public Property Get Markdown() As String
    Markdown = mMarkdown
End Property

Public Property Let Position(var As Long)
Attribute Position.VB_Description = "The position this block was opened."
    If var > 0 And mPosition = 0 Then mPosition = var
End Property

Public Property Get Position() As Long
    Position = mPosition
End Property

Public Property Get Priority() As Long
Attribute Priority.VB_Description = "The priority level for this block."
'   The base block type has the lowest priority.
    Priority = 0
End Property

Public Property Set Range(var As Range)
Attribute Range.VB_Description = "The location range in the document for this content."
    Set mRange = var
End Property

Public Property Get Range() As Range
    With mRange
        If .End <> .Start + Len(mText) Then
            Dim docLength As Long
            Dim rangeEnd As Long

            docLength = Len(.Parent.content)
            rangeEnd = .Start + Len(mText)

            Set mRange = .Parent.Range(.Start, Utils.Min(rangeEnd, docLength))
        End If
    End With
    Set Range = mRange
End Property

Public Property Let SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort number for this content."
    mSortIndex = var
End Property

Public Property Get SortIndex() As Long
    SortIndex = mSortIndex
End Property

Public Property Get Valid() As Boolean
Attribute Valid.VB_Description = "Whether this block is valid or not."
'   The base block type is always valid.
    Valid = True
End Property

' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Sub Close(truncatePosition As Long)
Attribute Close.VB_Description = "Triggers a forced block closure and truncates markdown."
'   Triggers a forced block closure and truncates markdown.
'   Closed and invalid blocks are garbage collected by the register.
'
'   Args:
'       truncatePosition: The position to truncate the markdown to.
'
'   Raises:
'       BlockBadArguments if the truncate position is greater than the markdown.
'
    If Len(mMarkdown) < truncatePosition Then Throw = Errs.BlockBadArguments
    mMarkdown = Left(mMarkdown, truncatePosition)
    mClosed = True
End Sub

Public Sub InjectBlockRegister(register As InlineBlockRegister)
Attribute InjectBlockRegister.VB_Description = "Injects the block register so that the block can interact directly."
'   Injects the block register so that the block can interact directly.
'
'   Args:
'       register: The block register to attach to.
'
    Set mRegister = register
End Sub

Public Sub ParseCharacter(char As String)
Attribute ParseCharacter.VB_Description = "Parses a character."
'   This sub should only control whether the block remains open or not.
'   If it is, the markdown is updated by the register.
'
'   Args:
'       char: The character to be parsed.
'
'   Raises:
'       BlockBadArguments if the string length greater than 1.
'
    If Len(char) > 1 Then Throw = Errs.BlockBadArguments
    If char = "" Then mRegister.CloseBlock Me
End Sub

Public Sub StyleContent()
Attribute StyleContent.VB_Description = "Styles content in the document."
    Throw = Errs.WarnNotImplementedException
End Sub

Public Sub WriteContent()
Attribute WriteContent.VB_Description = "Writes content to the document."
    Throw = Errs.WarnNotImplementedException
End Sub


' Helpers
'-------------------------------------------------------------------------------


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Let ISortable_SortIndex(var As Long)
Attribute SortIndex.VB_Description = "The sort index for this item."
    SortIndex = var
End Property

Public Property Get ISortable_SortIndex() As Long
    ISortable_SortIndex = SortIndex
End Property

Public Property Let IBlockInline_SortIndex(var As Long)
    SortIndex = var
End Property

Public Property Get IBlockInline_SortIndex() As Long
    IInlineContent_SortIndex = SortIndex
End Property

Public Property Get IBlockInline_Closed() As Boolean
Attribute Closed.VB_Description = "Whether this block is closed or not."
    IBlockInline_Closed = Closed
End Property

Public Property Let IBlockInline_Markdown(var As String)
Attribute Markdown.VB_Description = "The string content that was parsed."
    Markdown = var
End Property

Public Property Get IBlockInline_Markdown() As String
    IBlockInline_Markdown = Markdown
End Property

Public Property Let IBlockInline_Position(var As Long)
Attribute Position.VB_Description = "The position this block was opened."
    Position = var
End Property

Public Property Get IBlockInline_Position() As Long
    IBlockInline_Position = Position
End Property

Public Property Get IBlockInline_Priority() As Long
Attribute Priority.VB_Description = "The priority level for this block."
    IBlockInline_Priority = Priority
End Property

Public Property Set IBlockInline_Range(var As Range)
    Set Range = var
End Property

Public Property Get IBlockInline_Range() As Range
    Set IBlockInline_Range = Range
End Property

Public Property Let IBlockInline_SortIndex(var As Long)
    SortIndex = var
End Property

Public Property Get IBlockInline_SortIndex() As Long
    IBlockInline_SortIndex = SortIndex
End Property

Public Property Get IBlockInline_Valid() As Boolean
Attribute Valid.VB_Description = "Whether this block is valid or not."
    IBlockInline_Valid = Valid
End Property

Public Sub IBlockInline_Close(truncatePosition As Long)
Attribute Close.VB_Description = "Triggers a forced block closure and truncates markdown."
    Close  
End Sub

Public Sub IBlockInline_InjectBlockRegister(register As InlineBlockRegister)
Attribute InjectBlockRegister.VB_Description = "Injects the block register so that the block can interact directly."
    InjectBlockRegister register
End Sub

Public Sub IBlockInline_ParseCharacter(char As String)
Attribute ParseCharacter.VB_Description = "Parses a character."
    ParseCharacter char
End Sub

Public Sub IBlockInline_StyleContent()
Attribute StyleContent.VB_Description = "Styles content in the document."
    StyleContent
End Sub

Public Sub IBlockInline_WriteContent()
Attribute WriteContent.VB_Description = "Writes content to the document."
    WriteContent
End Sub