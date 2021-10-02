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


'   String manipulation.
'-------------------------------------------------------------------------------
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
    If Not Len(dirtyChars) = Len(cleanChars) Then Throw = UtilsBadArguments
    
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
        TinyLTrim = line
    Else
        TinyLTrim = Mid(line, i)
    End If
End Function


'   String interrogation.
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

Function StrIn(str As String, subStr As String) As Boolean
'   Helper for InStr for when you don't care where subStr is.
    StrIn = InStr(str, subStr) > 0
End Function

Public Function EndsWithNewLine(str As String) As Boolean
'   Tests the passed in string to see if it ends in a new line.
'
'   Args:
'       str: The string to be tested.
'
'   Returns:
'       True if str ends in a carriage return or line feed.

    Select Case Right(str, 1)
        Case Is = vbCr, vbLf
            EndsWithNewLine = True
    End Select
End Function

Public Function Regex(patternString As String, testString As String) As Object
'   Executes a regex returning any and all matches.
'   Funcionality can be moved to a dedicated class if more complexity required.
'
'   Args:
'       patternString: The pattern to evaluate the testString with.
'       testString: The string to be evaluated.
'
'   Returns:
'       The regex match object which supports .Count and default method .Item(n)
'
    Dim re As Object
    Dim mo As Object

    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = patternString
        .Global = True
        Set mo = .Execute(testString)
        Set Regex = mo
    End With

    Logger.Log "RegEx " & mo.Count & " matches " & _
        patternString & " with " & testString
End Function

'   Long Helpers
'-------------------------------------------------------------------------------
Public Function Between(n As Long, a As Long, b As Long) As Boolean
'   Returns True if n is between or equal to a and b
'
'   Args:
'       n: The number to test.
'       a: The lower bound.
'       b: The upper bound.
'
'   Returns:
'       True if n between or equal to a or b.
'
    Between = n >= a And n <= b   
End Function

Public Function Min(a As Long, b As Long) As Long
'   Returns the smaller of the two numbers.
'
'   Args:
'       a: The first number to compare.
'       b: The second number to compare.
'
'   Returns:
'       The smaller of a and b.
'
    Min = IIf(a < b, a, b)
End Function


'   Block helpers
'-------------------------------------------------------------------------------
Public Sub CBlockLeaf( _
    block As IBlockContent, _
    castToBlock As IBlockLeaf, _
    Optional parentBlock As IBlockContainer)
'   Casts a block of any type to an IBlockLeaf of type castToBlock.
'
'   Args:
'       block: The block to be cast.
'       castToBlock: The block to cast to.
'       parentBlock: The parent container so its child ref can be updated.
'
    Logger.Log "Casting " & TypeName(block) & " to " & TypeName(castToBlock)

    If TypeOf block Is BlockLeafHeading Then
        If Not TypeOf castToBlock Is BlockLeafHeading Then
            Dim hdg As BlockLeafHeading
            Set hdg = block
            If hdg.HeadingLevel > 0 Then
                Throw = Errs.UtilsWarnDestructiveCast
            End If
        End If
    End If

'   Copy across the leaf content.
    castToBlock.ContentText = block.ContentText

    If TypeOf block Is IBlockContainer Then
'       Check container has no children.
        Dim container As IBlockContainer
        Set container = block
        If container.Children.Count > 0 Then
            Logger.Log "Destructive cast. Children lost.", Level.Warning
        End If
    End If

'   Get the child index from parent.
    Dim childIndex As Long
    If Not parentBlock Is Nothing Then
        childIndex = parentBlock.Children.PopFind(block)
    End If

'   Cast to new block.
    Set block = castToBlock

'   Update the child reference on the parent.
    If Not parentBlock Is Nothing Then
        parentBlock.Children.Insert childIndex, block
    End If
End Sub

'   Document helpers
'-------------------------------------------------------------------------------
Public Sub WriteContent(toDoc As Document, block As IBlockContainer)
'   This method MUST be called before StyleContent.

'   Writes the document by recursively calling the WriteContent method
'   of all blocks in Children.
'
'   Args:
'       toDoc: The document to write to.
'       block: The block to be written.
'
    If TypeOf block Is IBlockContent Then
        Dim bCont As IBlockContent
        Set bCont = block
        If Not bCont.ContentText = vbNullString Then
            Logger.Log "BlockContainer contains content text!", _
                Level.Warning
            Logger.Log "Not written: '" & bCont.ContentText & "'", _
                Level.Warning
        End If
    End If

    Dim child As IBlock
    For Each child In block.Children
        child.WriteContent toDoc
    Next
End Sub

Public Sub StyleContent(styleDoc As Document, block As IBlockContainer)
'   This method MUST be called after WriteContent.

'   Styles the document by recursively calling the StyleContent method
'   of all blocks in Children.
'
'   Args:
'       styleDoc: The document to style.
'       block: The block to be styled.
'
    Dim child As IBlock
    Do While block.Children.Items.Count > 0
        Set child = block.Children.Pop
        child.StyleContent styleDoc
    Loop
End Sub
