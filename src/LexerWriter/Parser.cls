VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: Parser
'   Static functions that parse text without relying on state.
'   Lines should always be passed unindented unless otherwise stated.
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Function IsCodeBlockFence( _
    line As String, ByRef fence As String) As Boolean
'   Detects a code fence in the line.
'   The fence is either vbNullString or the current fence, indicating not in
'   or in a code block respectively. Fence is updated or cleared when detected.
'
'   Args:
'       line: The line of text to parse.
'       fence: The current code fence.
'
'   Returns:
'       True if a code block fence was detected. Also updates fence.
    Const VALID_FENCE_CHARS As String = "`~"

'   Remove non-indent level spacing and test minimum length.
    line = Utils.TinyLTrim(line)
    If Len(line) < 3 Then Exit Function

'   Mode 1: Currently in fence block.
    If Not fence = vbNullString Then
        If Left(line, Len(fence)) = fence Then
            fence = vbNullString
            IsCodeBlockFence = True
        End If
        Exit Function
    End If

'   Mode 2: Not currently in fence block.
    Dim countFence As Long
    Dim fenceChar  As String

    fenceChar = Left(line, 1)
    If Utils.StrIn(VALID_FENCE_CHARS, fenceChar) Then
'       Count contiguous fence characters.
        For countFence = 2 To Len(line)
            If Not Mid(line, countFence, 1) = fenceChar Then
                Exit For
            End If
        Next countFence

'       3 or more contiguous fence characters is a fence.
        countFence = countFence - 1
        If countFence > 2 Then
            fence = Left(line, countFence)
            IsCodeBlockFence = True
        End If
    End If
End Function