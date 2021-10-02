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


' Converters
'-------------------------------------------------------------------------------
Public Function BlockTypeToStyleName( _
    blType As BlockType, Optional modifier As Long = 0) As String
'   Converts a BlockType enum to a style LocalName.
'
'   Args:
'       blType: The BlockType to convert.
'       modifier: Currently only used for heading levels.
'
'   Returns:
'       A string matching the LocalName property of a style.
'
    Dim localName As String
    Select Case blType
        Case Is = BlockType.Paragraph
            localName = "Normal"
        Case Is = BlockType.Heading
            If modifier = 0 Then
                localName = "Title"
            Else
                localName = "Heading " & modifier
            End If
        Case Is = BlockType.FencedCode, BlockType.IndentedCode
            localName = "CodeBlock"
        Case Else
            Logger.Log "No explicit style for BlockType(" & blType & ")", _
                Level.Information
            localName = "Normal"
    End Select

    BlockTypeToStyleName = localName
End Function
