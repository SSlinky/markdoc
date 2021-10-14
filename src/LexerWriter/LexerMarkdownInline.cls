VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LexerMarkdownInline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: LexerMarkdownInline
'   A lexer that focuses on a single block.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
'   TODO: ISSUE This pattern doesn't work for tick escaped ticks inside code.
'           e.g. Not code `this is ````` code`.
'           The ````` should be parsed as is and interpreted as inline code.
Private Const INLINECODEPATTERN As String = "((?:^|(?:[^\\`]))(?:\\\\)*)((?:`)((?:[^`\n]+?))(?:`))"
Private Const ESCAPABLE As String = "\!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"

Private mBlock As IBlockContentInline
Private mMatchDict As Object
Private mThisInline As InlineContent
Private mInlineContents As List
Private mCurrentLineStart As Long


' Properties
'-------------------------------------------------------------------------------


' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Sub Parse(block As IBlockContentInline, blockParent As IBlockContainer)
Attribute ParseInlineContent.VB_Description = "Parses the inline content of the block."
'   Parses the inline content of the block.
'
'   Args:
'       block: The block that has inline content.
'
    Set mBlock = block
    Set mInlineContents = New List
'
'   Prepare the content variables.
    Dim line As Variant
    Dim lines As List
    Dim thisLine As String
    Dim prevLine As String
    Dim splitLines As Variant

    Set lines = New List
    splitLines = Split(block.ContentText, vbNewLine)
    lines.FromArray splitLines

'   Prepare the parser variables.
    Dim i As Long
    Dim c As String * 1

'   Prepare parser state flags.
    Dim isCode As Boolean
    Dim isBold As Boolean
    Dim isItalic As Boolean
    Dim isStruck As Boolean
    Dim isEscaped As Boolean

    Set mThisInline = New InlineContent

    For Each line In lines    
        prevLine = thisLine
        thisLine = PrepLine(prevLine, line)

'       Break out if we hit a line break.
        If Parser.HasHardLineBreak(thisLine, prevLine) Then
            SplitBlockLeaf block, blockParent, _
                Mid(block.ContentText, mCurrentLineStart)
            Exit For
        End If

'       Parse content with Regex
        Set mMatchDict = CreateObject("Scripting.Dictionary")

        thisLine = ParseWithRegex INLINECODEPATTERN, thisLine, _
            DocumentShortcuts.GetStyle("CodeInline")

'       Iterate over each character not already parsed.
        For i = 1 To Len(thisLine)
'           Set the sort index if it's a new block.
            If mThisInline.SortIndex = 0 Then
                mThisInline.SortIndex = i + mCurrentLineStart
            End If

'           Ignore sections already parsed with Regex
            i = SkipMatchedSections(i)

            c = Mid(thisLine, i, 1)
            Select Case True
'               Handle erased section.
                Case Is = c = Chr(0)

'               Handle an escape character.
                Case Is = Not isEscaped And c = "\"
                    isEscaped = True

'               Handle the character after the escape character.
                Case Is = isEscaped
                    If Not Utils.StrIn(ESCAPABLE, c) Then c = "\" & c
                    mThisInline.AppendText c
                    isEscaped = False

'               Handle plain text.
                Case Else
                    mThisInline.AppendText c

            End Select
        Next i
        LazyPushThisInline
    Next line

'   Apply all to block.
    Set block.InlineContents = mInlineContents
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Function ParseWithRegex(pattern As String, thisLine As String, sty As Style) As String
Attribute ParseWithRegex.VB_Description = "Updates mMatchDict with match details."
'   Updates mMatchDict with match details. The dictionary is an object so
'   is passed byRef which means we do not need to explicitly set a return.
'
'   Args:
'       pattern: The regex pattern.
'       thisLine: The line to test.
'       sty: The style to apply.
'
'   Returns:
'       The original line that was tested but any matched sections erased.
'
'   Due to the fact that only the entire match object gives us a FirstIndex
'   property, the below three capture groups are required. This is to ensure the
'   right sections are captured / replaced.
'
'   Group 1: The section immediately before Group 2.
'   Group 2: The section of thisLine to be replaced with null characters.
'            This section will also be flagged to be skipped by the parser.
'   Group 3: The section to be stored in the Text property of InlineContent.
'
'   Groups should be nested in the following format.
'   (1)(2(3))
'
    Dim regexResult As Object
    Dim matchObject As Object
    Dim grpOneFirstIndex As Long
    Dim grpTwoFirstIndex As Long
    Dim grpTwoLastIndex As Long

    Set regexResult = Utils.Regex(INLINECODEPATTERN, thisLine)

    For Each matchObject In regexResult
'       Get the match object metadata.
        grpOneFirstIndex = matchObject.FirstIndex + 1
        grpTwoFirstIndex = grpOneFirstIndex + Len(matchObject.SubMatches(0))
        grpTwoLastIndex = grpTwoFirstIndex + Len(matchObject.SubMatches(1)) - 1

'       Add a skip item to the dictionary.
'           Key: Position start.
'           Value: Position after skip.
        mMatchDict.Add grpTwoFirstIndex + 1, grpTwoLastIndex

'       Set the values on the InlineContent.
        mThisInline.Text = matchObject.SubMatches(2)
        mThisInline.SortIndex = mCurrentLineStart + grpOneFirstIndex
        Set mThisInline.Style = sty

'       Erase the matched section of the string so it cannot be parsed with another regex.
        thisLine = Utils.EraseStringSection(thisLine, grpTwoFirstIndex, grpTwoLastIndex)

'       Push the InlineContent and reset.
        LazyPushThisInline
    Next matchObject

'   Return the censored line.
    ParseWithRegex = thisLine
End Function

Private Sub SplitBlockLeaf(block As IBlockLeaf, parent As IBlockContainer, splitContent As String)
Attribute SplitBlock.VB_Description = "Splits content of one block over two."
'   Splits content of one block over two with a blank line between.
'
'   Args:
'       block: The block to split.
'       parent: The parent block.
'       splitContent: The ContentText for the split block.
'
    Dim insertAtIndex As Long
    insertAtIndex = parent.Children.IndexOf(block) + 1

    Dim splitBlock As IBlockLeaf
    Set splitBlock = BlockFactory.CreateBlock(TypeName(block))
    splitBlock.ContentText = splitContent

    parent.Children.Insert insertAtIndex, splitBlock
    parent.Children.Insert insertAtIndex, New BlockLeafBlankLine
End Sub

Private Function SkipMatchedSections(i As Long) As Long
Attribute SkipMatchedSections.VB_Description = "Skips over sections matched with Regex."
'   Skips over sections matched with Regex. Sections the InlineContent if required.
'
'   Args:
'       i: The current line iterator position.
'
'   Returns:
'       The iterator line position after content skipped.
'
'   Raises:
'       ApplicationValueError if the skipped section will cause endless loop.
'
    Dim skipLength As Long

    Do While mMatchDict.Exists(i + 1)
        skipLength = mMatchDict.Item(i + 1)
        If skipLength <= i Then Throw = Errs.ApplicationValueError

        LazyPushThisInline
        i = skipLength
        mThisInline.SortIndex = i + mCurrentLineStart
    Loop
    SkipMatchedSections = i
End Function

Private Sub LazyPushThisInline()
Attribute LazyPushThisInline.VB_Description = "Pushes the current InlineContent if it has text."
'   Pushes the current InlineContent if it has text
'
    If Not mThisInline.Text = vbNullString Then
        mInlineConts.PushSort mThisInline

        With mThisInline
            Logger.Log "PushSort " & .SortIndex & " - " & .Text
        End With

        Set mThisInline = New InlineContent
    End If
End Sub

Private Function PrepLine(prevLine As String, thisLine As Variant) As String
Attribute PrepLine.VB_Description = "Casts to string and prepends with a space if required."
'   Casts to string and prepends with a space if required
'
'   Args:
'       prevLine: The previous line (used to detirmine if spacing required)
'       thisLine: The current line to be cast.
'
'   Returns:
'       A string line prepended with a space if previous line has text.
'
    PrepLine = IIf(prevLine = vbNullString, "", " ") & CStr(thisLine)
End Function
