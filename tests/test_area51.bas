Attribute VB_Name = "test_area51"
Option Explicit

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


Function RegexTest() As Object
'    Const PAT As String = "^\s{,3}(?P<lvl>#{1,6})\s+(?P<heading>.*?)\s*(?:#*\s*)?$"
'    Const PAT As String = "^\s*#*\s+(.*?)\s*(?:#*\s*)?$" ' works!
'    Const PAT As String = "^\s{,3}(#{1,6})\s+(.*?)\s*(?:#*\s*)?$" ' works!
    Const PAT As String = "^\s{0,3}(#{1,6})\s+(.*?)\s*(?:#*\s*)?$"
    Dim t1 As String

    t1 = "    # Simple Heading"

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = PAT

    Dim m As Object
    Dim x As Object
    Dim s As Long
    Dim a As Long

    Set m = re.Execute(t1)
    Debug.Print "Matching " & t1
    For Each x In m
        Debug.Print x.Value
        a = x.Submatches.Count - 1
        For s = 0 To a
            Debug.Print "    " & s & ": " & x.Submatches(s)
        Next s
    Next x

    If m.Count Then Debug.Print "matched!"
    Set RegexTest = m

End Function

Public Function IsHeading(line As String, matches As Object) As Boolean
'   Detects a heading.
'
'   Args:
'       line: The line of text to parse.
'       matches: The regexp match object to save calling this multiple times.
'
'   Returns:
'
'   Raises:
'       True if a code block fence was detected. Also updates fence.

    Const HEADING_PATTERN As String = "^\s{0,3}(#{1,6})\s+(.*?)\s*(?:#*\s*)$"
    Set matches = Regex(HEADING_PATTERN, line)
    IsHeading = matches.Count > 0
End Function

Public Function Regex(patternString As String, testString As String) As Object
'   Executes a regex returning matches if any.
'
'   Args:
'       pattern:
'       toTest:
'
'   Returns:
'       The regex match object.

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = patternString
        Set Regex = .Execute(testString)
    End With
End Function

Public Sub test_IsHeading()

    Dim matches As Object

    If IsHeading("   ###### Simple Heading   ########   x ####", matches) Then
        With matches(0)
            Debug.Print .Submatches(0), .Submatches(1)
        End With
    End If
End Sub



Sub test_Cast()

    Dim b As IBlock
    Dim x As IBlock

    Set b = New BlockContainer
    Set x = b

    Utils.CBlockLeaf b, New BlockLeafHeading
    Debug.Print TypeName(b)
    Debug.Print "Is BlockLeafHeading: " & TypeOf b Is BlockLeafHeading
    Debug.Print "Reference match: " & (b Is x) & vbNewLine

'    Dim bl As New List
'    bl.Push New BlockContainer
'    Debug.Print TypeName(bl.Peek)
'
'    Utils.CBlockLeaf bl.Peek, New BlockLeafHeading
'    Debug.Print TypeName(bl.Peek)
'    Debug.Print "Is BlockLeafHeading: " & TypeOf bl.Peek Is BlockLeafHeading
'
'    Utils.CBlockLeaf bl.Items(1), New BlockLeafHeading
'    Debug.Print "Is BlockLeafHeading: " & TypeOf bl.Items(1) Is BlockLeafHeading

    Dim col As New Collection
    Set b = New BlockContainer
    col.Add b

    Utils.CBlockLeaf col.Item(1), New BlockLeafHeading
    Debug.Print TypeName(col.Item(1))
    Debug.Print "Collection item is BlockLeafHeading: " & TypeOf col.Item(1) Is BlockLeafHeading
    Debug.Print "b item is BlockLeafHeading: " & TypeOf b Is BlockLeafHeading
    Debug.Print "Reference match: " & (b Is col.Item(1))
End Sub

Sub Cast(x As IBlock, toX As IBlockLeaf)
    Set x = toX
End Sub

Function GetIBlock(x As IBlock) As IBlock
    Set GetIBlock = x
End Function

Sub test_WarningLevel()
    Logger.Log "Default message level."
    Throw.Exception = Errs.FileReaderWarnEmptyFile
End Sub

Sub test_ISortable()

    Dim x As ISortable
    Dim xs As New List
    
    Dim nums() As String
    Dim printNums As String
    
    nums = Split("1,2,3,7,54,2,8,3,1,5,8,90,2,4,6", ",")
    Dim i As Long
    
    For i = 0 To UBound(nums)
        Set x = New InlineContent
        x.SortIndex = Int(nums(i))
        xs.PushSort x
        
        printNums = ""
        For Each x In xs
            printNums = printNums & x.SortIndex & " "
        Next x
        Debug.Print printNums
    Next i

    Debug.Assert printNums = "1 1 2 2 2 3 3 4 5 6 7 8 8 54 90 "
End Sub
