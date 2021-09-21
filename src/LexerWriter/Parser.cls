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
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Function IsCodeBlockFence(line As String, ByRef fence As String) As Boolean
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

End Sub