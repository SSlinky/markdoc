VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Static Class: Utils
'   Contains generic static helper functions.
'-------------------------------------------------------------------------------

Const ASCII_NBSP As Long = 160
Const ASCII_TAB As Long = 9


'   String manipulation and interrogation.
'-------------------------------------------------------------------------------
Public Function CountSubstr(str As String, subStr As String) As Long
'   Counts the number of times subStr is present in str.
    CountSubstr = (Len(str) - Len(Replace(str, subStr, ""))) / Len(subStr)
End Function

Public Function GetIndentationLevel(line As String) As Integer
'   Gets indent by counting spaces preceding any non-space character.
    line = Replace(line, Chr(ASCII_TAB), "    ")
    If Trim(line) = "" Then Exit Function

'   Append a character so we can trim to find left indentation.
    GetIndentationLevel = Fix((Len(line) - Len(Trim(line & ".")) + 1) / 4)
End Function

Public Function UnindentLine(line As String, indentLevel As Long) As String
'   Removes indentation from a line so that it can be parsed as if it was
'   at base indent level.
'
'   Args:
'       line: The text to remove indentation from.
'       indentLevel: The level of indentation to remove.
'
'   Returns:
'       The string with the indentation removed.
'
'   Raises:
'       UtilsStringIndentNotBlank when the indented section not tabs or spaces.
'       UtilsBadArguments when indentLevel below zero or longer than line.

'   Don't do work if there's no work.
    If indentLevel = 0 Then
        UnindentLine = line
        Exit Function
    End If

'   Throw if there's no line or the indent level is less than 0.
    If line = vbNullString or indentLevel < 0 Then
        Throw = Errs.UtilsBadArguments
        UnindentLine = line
        Exit Function
    End If

    dim i As Long
    dim c As String
    
    For i = 1 To indentLevel
'       Throw if we hit the end of the line.    
        If i > Len(line) Then
            Throw = Errs.UtilsBadArguments
            UnindentLine = line
            Exit Function
        End If

        c = Mid(line, i, 1)
        Select Case c
'           Iterate only when we hit a space.        
            Case Is = " ":

'           Replace tabs with four spaces.            
            Case Is = Chr(ASCII_TAB):
                line = InsertReplace(line, "    ", i, 1)

'           Throw for any other character.                
            Case Else:
                Throw = Errs.UtilsStringIndentNotBlank
                UnindentLine = line
                Exit Function
        End Select
    Next i

    UnindentLine = Mid(line, i)
End Function

Public Function InsertReplace( _
    str As String, subStr As String, pos As Long, length As Long) As String
'   Inserts a substring into a string replacing a section starting
'   at pos for length characters. This function updates your pos and length
'   arguments if they will cause an exception to be thrown.
'
'   Args:
'       str: The string to perform the replacement on.
'       subStr: The substring to insert.
'       pos: The position start to replace.
'       length: The position length to replace.
'
'   Returns:
'       str with subStr replace inserted into the correct place.

    If length < 0 Then
        pos = pos + length
        length = Abs(length)
    End If

    If pos < 0 Then
        length = length + pos
        pos = 0
        If length < 0 Then length = 0
    End If

    Select Case True
        Case Is = pos = 0 And length = 0:
            InsertReplace = subStr & str
        Case Is = pos >= Len(str):
            InsertReplace = str & subStr
        Case Is = pos = 0:
            InsertReplace = subStr & Mid(str, length)
        Case Else:
            InsertReplace = Left(str, pos - 1) & subStr & Mid(str, pos + length)
    End Select
End Function

Function SuperTrim(str As String, Optional trimChars As String) As String
'   Removes any character in trimChars from the start or end of the line.
    Dim i As Long
    
'   Default trim characters
    If trimChars = "" Then
        trimChars = vbNewLine & Chr(ASCII_NBSP) & Chr(ASCII_TAB) & " "
    End If
    
'   Trim left
    For i = 1 To Len(str)
        If Not StrIn(Mid(str, i, 1), trimChars) Then Exit For
    Next i
    If i > 1 Then str = Right(str, Len(str) - i + 1)

'   Trim right
    For i = Len(str) To 1 Step -1
        If Not StrIn(Mid(str, i, 1), trimChars) Then Exit For
    Next i
    If i < Len(str) Then str = Left(str, i)
    
    SuperTrim = str
End Function

Function CleanString(str As String, _
    Optional dirtyChars As String, _
    Optional cleanChars As String) As String

'   Converts any single dirtyChar into a cleanChar.
'   dirtyChars and cleanChars must be the same length.
    If Len(dirtyChars) = Len(cleanChars) Then Throw = UtilsBadArguments
    
'   Set up defaults if nothing passed
    If dirtyChars = "" Then
        dirtyChars = Chr(ASCII_NBSP)
        cleanChars = " "
    End If
    
    Dim i As Long
'   Replace each dirtyChar with the corresponding cleanChar
    For i = 1 To Len(dirtyChars)
        str = Replace(str, Mid(dirtyChars, i, 1), Mid(cleanChars, i, 1))
    Next i
    CleanString = str
End Function

Function StrIn(str As String, subStr As String) As Boolean
'   Helper for InStr for when you don't care where subStr is.
    StrIn = InStr(str, subStr) > 0
End Function

Function TinyLTrim(line As String) As String
'   Removes up to the first three spaces from the left of line.
'
'   Args:
'       line: The line of text to trim.
'
'   Returns:
'       The line, minus up to three spaces on from the left side.

    Dim i As Long
    For i = 1 To 3
        If Mid(line, i, 1) <> " " Then
            Exit For
        End If
    Next i

    If i = 1 Then
        TinyTrim = line
    Else
        TinyTrim = Mid(line, i)
    End If
End Function


'   Long Helpers
'-------------------------------------------------------------------------------
Public Function LBetween(n As Long, a As Long, b As Long) As Boolean
'   Returns True if n is between or equal to a and b
'
'   Args:
'       n: The number to test.
'       b: The lower bound.
'       c: The upper bound.
'
'   Returns:
'       True if n between or equal to a or b.
'
    LBetween = n >= a And n <= b   
End Function


'   Block helpers
'-------------------------------------------------------------------------------
Public Sub CBlockLeaf(block As IBlock, castToBlock As IBlockLeaf)
'   Casts a block of any type to an IBlockLeaf of type castToBlock.
'   This may cause loss of data as paragraphs do not have children.

    If TypeOf block Is IBlockLeaf Then
'       Copy across the leaf content.
        Dim leaf As IBlockLeaf
        Set leaf = block
        castToBlock.Content = leaf.ContentText
        Set block = castToBlock
    ElseIf TypeOf block Is IBlockContainer Then
'       Containers have no content.
        Set block = castToBlock
    Else
        Throw = Errs.BlockTypeMismatch
    End If
End Sub
