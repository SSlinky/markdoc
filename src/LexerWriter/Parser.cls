VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("markDoc.src.LexerWriter")
Option Explicit

'-------------------------------------------------------------------------------
'   Static Class: Parser
'   Functions that parse text without relying on state.
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
    Dim countFence As Long
    Dim fenceChar  As String

    If Not fence = vbNullString Then
'       Check fence matches left side of line
        If Left(line, Len(fence)) = fence Then
'           Check entire line only fence characters.
            fenceChar = Left(fence, 1)
            countFence = Utils.CountSubstr(Trim(line), fenceChar)
            If countFence = Len(Trim(line)) Then
                fence = vbNullString
                IsCodeBlockFence = True
            End If
        End If
        Exit Function
    End If

'   Mode 2: Not currently in fence block.
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

Public Function IsHeading(line As String, matches As Object) As Boolean
'   Detects a heading.
'
'   Args:
'       line: The line of text to parse.
'       matches: The regexp match object passed byRef to save multiple calls.
'
'   Returns:
'
'   Raises:
'       True if a code block fence was detected. Also updates fence.

    Const HEADING_PATTERN As String = "^\s{0,3}(#{1,6})\s+(.*?)\s*(?:#*\s*)$"
    Set matches = Utils.Regex(HEADING_PATTERN, line)
    IsHeading = matches.Count > 0
End Function

Public Function IsListItem(line As String, matches As Object) As Boolean
'   Detects a list item.
'
'   Args:
'       line: The line of text to parse.
'
'   Returns:
'
'   Raises:
'
    Throw = Errs.WarnNotImplementedException
End Function

Public Function ParseUnknownBlockLine(line As String, matches As Object) As BlockType
'   Parses a line as is without any assumptions.
'   This line is typically the beginning of a new block.
'
'   Args:
'       line: The line to be parsed.
'       matches: The regexp match object passed byRef to save multiple calls.
'
'   Returns:
'       The determined block type.

    Select Case True
        Case Is = IsHeading(line, matches)
            ParseUnknownBlockLine = BlockType.Heading
            Logger.Log "Parse resolved to: Heading"
        Case Else
            ParseUnknownBlockLine = BlockType.Paragraph
            Logger.Log "Parse resolved to: Paragraph"
    End Select
End Function

Public Function ParseHeadingLevel(headingText As String) As Long
'   Determines the heading level from the raw markdown text.
'   The passed string is assumed to be a valid heading, therefore no checks
'   are done so results may be unexpected if passing an unvalidated string.
'
'   Args:
'       headingText: The raw markdown text that describes the heading.
'
'   Returns:
'       The length of the first string of non-space characters.

    ParseHeadingLevel = Len(Split(Trim(headingText), " ")(0)) - 1
End Function

Public Function HasHardLineBreak(line As String, prevLine As String) As Boolean
Attribute HasHardLineBreak.VB_Description = "Test line for hard line break."
'   Test line for hard line break in a paragraph line.
'
'   Args:
'       line: The line of text to be parsed.
'       prevLine: The previous line of parsed text.
'
'   Returns:
'       True if the line should be broken.
'
    If line = vbNewLine Or prevLine = vbNullString Then Exit Function
    Const DOUBLESPACE As String = "  "

    If Utils.EndsWith(prevLine, DOUBLESPACE) Then
        HasHardLineBreak = True
        Exit Function
    End If
End Function

Public Function LeftFlankDelimiterRun(str As String, ind As Long) As String
Attribute LeftFlankDelimiterRun.VB_Description = "Returns the left flank of a delimiter run."
'   Returns the left flank of a delimiter run.
'
'   Args:
'       str: The string to test for a delimiter run.
'       ind: The start (left) position of the delimiter.
'
'   Returns:
'       A string delimiter if one exists or null string.
'
'   A left-flanking delimiter run is a delimiter run that is
'       (1) not followed by Unicode whitespace, and either
'       (2a) not followed by a Unicode punctuation character, or;
'       (2b1) followed by a Unicode punctuation character and
'       (2b2) preceded by Unicode whitespace or a Unicode punctuation character.
'   For purposes of this definition, the beginning and the end of the line count as Unicode whitespace.
'
    LeftFlankDelimiterRun = FlankDelimiterRun(str, ind, 1)
End Function

Public Function RightFlankDelimiterRun(str As String, ind As Long) As String
Attribute LeftFlankDelimiterRun.VB_Description = "Returns the right flank of a delimiter run."
'   Returns the right flank of a delimiter run.
'
'   Args:
'       str: The string to test for a delimiter run.
'       ind: The start (left) position of the delimiter.
'
'   Returns:
'       A string delimiter if one exists or null string.
'
'   A right-flanking delimiter run is a delimiter run that is
'       (1) not preceded by Unicode whitespace, and either
'       (2a) not preceded by a Unicode punctuation character, or;
'       (2b1) preceded by a Unicode punctuation character and
'       (2b2) followed by Unicode whitespace or a Unicode punctuation character.
'   For purposes of this definition, the beginning and the end of the line count as Unicode whitespace.
'
    RightFlankDelimiterRun = FlankDelimiterRun(str, ind, 1)
End Function

Private Function FlankDelimiterRun(str As String, ind As Long, lro As Long) As String
Attribute FlankDelimiterRun.VB_Description = "Performs the work of finding a left or right delimiter run."
'   Performs the work of finding a left or right delimiter run.
'   This function is called by Left or Right delimiter run functions.
'
'   Args:
'       str: The string to test for a delimiter run.
'       ind: The start (left) position of the potential delimiter.
'       lro: Left(1)/Right(-1) offset.
'
'   Returns:
'       A string delimiter if one exists or null string.
'
'   Raises:
'       ParserBadArguments if the left-right offset is not valid.
'
    If Not lro ^ 2 = 1 Then Throw = Errs.ParserBadArguments

    Dim c As String * 1
    Dim pc As String * 1
    Dim nc As String * 1
    Dim rtn As String
    
    c = Mid(str, ind, 1)
    
'   Strong and emphasis only use _ or *
    If c <> "*" And c <> "_" Then Exit Function
    
    str = " " & str & " "   ' Pad the string with spaces.
    pc = Mid(str, ind)      ' Set the previous char based on current index.
    ind = ind + 1           ' Offset the index to account for padding.
    
'   Test for double (strong) or single (emphasis)    
    If Mid(str, ind + 1, 1) = c Then
        rtn = c & c
        ind = ind + lro
    Else
        rtn = c
    End If
    
'   Set the next char based on the lro direction.    
    nc = Mid(str, ind + lro, 1)
    
'   Previous and next char punctuation and whitespace flags.
    Dim ncIsWhitespace As Boolean
    Dim pcIsWhitespace As Boolean
    Dim ncIsPunctuation As Boolean
    Dim pcIsPunctuationOrWs As Boolean
    
    ncIsWhitespace = nc = " " Or nc = Chr(160)
    pcIsWhitespace = pc = " " Or pc = Chr(160)
    
'   The next character cannot be whitespace.    
    If Not ncIsWhitespace Then
        ncIsPunctuation = IsUnicodePunctuation(nc)
        pcIsPunctuationOrWs = pcIsWhitespace Or IsUnicodePunctuation(pc)

'       The next character is not punctuation, or if it is, the previous
'       character is also punctuation or whitespace.
        If Not ncIsPunctuation Or (ncIsPunctuation And pcIsPunctuationOrWs) Then
            FlankDelimiterRun = rtn
        End If
    End If   
End Function
