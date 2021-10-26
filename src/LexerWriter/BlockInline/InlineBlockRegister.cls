VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineBlockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.LexerWriter.BlockInline")
Option Explicit
Public Event OnOpenBlock(block As IBlockInline)
Public Event OnCloseBlock(block As IBlockInline)
Public Event OnParse(char As String, pos As Long)
'-------------------------------------------------------------------------------
'   Class: InlineBlockRegister
'   Manages active inline blocks.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mItems As List
Private mMarkdown As String

Private mBacktickCount As Long
Private mBlockInlineCodeOpen As Boolean

' Properties
'-------------------------------------------------------------------------------
Public Property Set Items(var As List)
Attribute Items.VB_Description = "The backing store for the register."
    Set mItems = var
End Property

Public Property Get Items() As List
    Set Items = mItems
End Property

Public Property Let Markdown(var As String)
Attribute Markdown.VB_Description = "The markdown this register is related to."
    If mMarkdown = vbNullString Then mMarkdown = var
End Property

Public Property Get Markdown() As String
    Markdown = mMarkdown
End Property


' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Sub ParseCharacter(pos As Long)
Attribute ParseCharacter.VB_Description = "Parses the character and creates a block if required."
'   Parses the character and creates a block if required.
'
'   Args:
'       pos: The character position in the markdown to parse.
'
    Dim char As String
    char = Mid(Markdown, pos, 1)

'   Each registered block should parse the character.
    RaiseEvent OnParse(char, pos)

'   Open a new block if we need to.
    LazyParseOpenBlock char, pos
End Sub

Public Sub Register(block As IBlockInline)
Attribute Register.VB_Description = "Registers a block for tracking."
'   Registers a block for tracking.
'
'   Args:
'       block: The IBlockInline to register.
'
    mItems.Push block
    RaiseEvent OnOpenBlock(block)
End Sub

Public Sub Unregister(block As IBlockInline)
Attribute Register.VB_Description = "Removes a block from tracking."
'   Removes a block from tracking.
'
'   Args:
'       block: The IBlockInline to unregister.
'
    Dim i As Long
    i = mItems.IndexOf(block)
    If i > 0 Then mItems.Remove i
End Sub

Public Sub CloseBlock(block As IBlockInline)
Attribute CloseBlock.VB_Description = "Raises the event that a block is closed."
'   Raises the event that a block is closed.
'
'   Args:
'       block: The IBlockInline to close.
'
    RaiseEvent OnCloseBlock(block)
End Sub

Public Sub BlockInlineCodeOpen()
Attribute BlockInlineCodeOpen.VB_Description = "Prevents inline code being opened on the same tick."
'   Prevents inline code being opened on the same tick.
'   This should be called by a code block when it is closed to stop the
'   same back ticks opening another block.
'
    mBlockInlineCodeOpen = True
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub LazyParseOpenBlock(char As String, pos As Long)
Attribute LazyParseOpenBlock.VB_Description = "Pushes a new parallel inline block if char is the trigger for any type."
'   Pushes a new parallel inline block if char is the trigger for any type.
'
'   Args:
'       char: a single character that could start a new inline block.
'       pos: The character position in the markdown to parse.
'
    Const TRIGGERS As String = "<`*_[!"
    If Not Utils.StrIn(TRIGGERS, char) Then Exit Sub

    Dim b As IBlockInline

'   Open a code block if we had back ticks and now we don't.
    If Not char = "`" And mBacktickCount > 0 Then
        If Not mBlockInlineCodeOpen Then
            Set b = New BlockInline
            b.Text = String("`", mBacktickCount)
            mItems.RegisterSelf(Me, pos - mBacktickCount)
        End If
        mBacktickCount = 0
    End If
    mBlockInlineCodeOpen = False

    Select Case char
        Case Is = "<"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
        Case Is = "`"
            mBacktickCount = mBacktickCount + 1
        Case Is = "*", "_"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
        Case Is = "["
            Set b = New BlockInline ' TODO: Update to correct concrete type.
        Case Is = "!"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
        Case Else
'           Drop out if nothing interesting discovered.
            Exit Sub
    End Select

'   Set the text and position and register.
    b.Text = char
    mItems.RegisterSelf(Me, pos)
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------

