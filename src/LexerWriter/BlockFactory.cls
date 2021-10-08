VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: BlockFactory
'   Generates blocks by their name.
'-------------------------------------------------------------------------------

' Methods
'-------------------------------------------------------------------------------
Public Function CreateBlock(blockName As String) As IBlock
Attribute CreateBlock.VB_Description = "Returns a block of type blockName."
'   Returns a block of type blockName
'
'   Args:
'       blockName: The type name of the block to be generated.
'
'   Returns:
'       An IBlock of type blockName.
'
'   Raises:
'       Bad argument if block type not registered.
'   
    Select Case blockName
        Case Is = "BlockLeafParagraph"
            Set CreateBlock = New BlockLeafParagraph
        Case Is = "BlockLeafHeading"
            Set CreateBlock = New BlockLeafHeading
        Case Is = "BlockLeafBlankLine"
            Set CreateBlock = New BlockLeafBlankLine
        Case Is = "BlockLeafFencedCode"
            Set CreateBlock = New BlockLeafFencedCode
        Case Is = "BlockLeafIndentedCode"
            Set CreateBlock = New BlockLeafIndentedCode
        Case Else
            Throw = Errs.UtilsBadArguments
    End Select
End Function
