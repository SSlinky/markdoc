Attribute VB_Name = "Enums"
Option Explicit

Public Enum BlockType
'   Containers
    List
    ListItem
    Quote

'   Leaves
    Heading
    BlankLine
    Paragraph
    FencedCode
    IndentedCode
End Enum

Public Function BlockTypeToStyleName(blType As BlockType) As String
'   Converts a BlockType enum to a style LocalName.
'
'   Args:
'       blType: The BlockType to convert.
'
'   Returns:
'       A string matching the LocalName property of a style.
'
    Dim localName As String
    Select Case blType
        Case Is = BlockType.Heading
            localName = "Heading "
        Case Is = BlockType.FencedCode, BlockType.IndentedCode
            localName = "CodeBlock"
        Case Else
            localName = "Normal"
    End Select

    BlockTypeToStyleName = localName
End Function
