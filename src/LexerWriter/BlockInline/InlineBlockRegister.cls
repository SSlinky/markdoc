VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineBlockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
'   TODO: Register does it's own parse to check for new blocks to be created.
End Sub

Public Sub Register(block As IBlockInline)
Attribute Register.VB_Description = "Registers a block for tracking."
'   Registers a block for tracking.
'
'   Args:
'       block: The IBlockInline to register.
'
    mItems.Push block
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


' Interface Implementation
'-------------------------------------------------------------------------------

