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
Public Event OnRegister(block As IBlockInline)

'-------------------------------------------------------------------------------
'   Class: InlineBlockRegister
'   Manages active inline blocks.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mItems As List


' Properties
'-------------------------------------------------------------------------------
Public Property Set Items(var As List)
Attribute Items.VB_Description = "The backing store for the register."
    Set mItems = var
End Property

Public Property Get Items() As List
    Set Items = mItems
End Property



' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Sub ParseCharacter(char As String)
Attribute ParseCharacter.VB_Description = "Parses the character and creates a block if required."
'   Parses the character and creates a block if required.
'
'   Args:
'       char: The character to parse.
'
'   Each registered block should parse the character.
    Dim block As IBlockInline
    For Each block In mItems
        block.ParseCharacter char
    Next block

'   TODO: Check for any closed blocks (first in priority). Push valid blocks
'         in the right order to the attached IBlockContent and bin invalid ones.

    LazyParseBlockFactory char
End Sub

Public Sub Register(block As IBlockInline)
Attribute Register.VB_Description = "Registers a block for tracking."
'   Registers a block for tracking.
'
'   Args:
'       block: The IBlockInline to register.
'
    mItems.Push block
    RaiseEvent OnRegister(block)
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


' Helpers
'-------------------------------------------------------------------------------
Private Sub LazyParseBlockFactory(char As String)
Attribute LazyParseBlockFactory.VB_Description = "Pushes a new parallel inline block if char is the trigger for any type."
'   Pushes a new parallel inline block if char is the trigger for any type.
'
'   Args:
'       char: a single character that could start a new inline block.
'
' TODO: move this to the class that actually parses. This class shouldn't know or care about this.
    Const TRIGGERS As String = "<`*_[!"
    If Not Utils.StrIn(TRIGGERS, char) Then Exit Sub

    Dim b As IBlockInline
    Select Case char
        Case Is = "<"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
            b.Text = char
        Case Is = "`"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
            b.Text = char
        Case Is = "*", "_"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
            b.Text = char
        Case Is = "["
            Set b = New BlockInline ' TODO: Update to correct concrete type.
            b.Text = char
        Case Is = "!"
            Set b = New BlockInline ' TODO: Update to correct concrete type.
            b.Text = char

    End Select
    mItems.Push b
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------

