Attribute VB_Name = "test_area51"
Option Explicit

Sub fhawl()
    Dim o As New List
    Dim r As Range
    Dim x As Range
    
    With ActiveDocument
        Set r = .Range(0)
        Set x = .Range(Len(.Content) - 1)
        r = "content"
        x = "xContentx"
    End With
    
    o.Push "abc"
    o.Push r
    
    Set r = Nothing
    
    o.Items.Remove 1
    Set r = o.Pop
    Debug.Print r.Text
End Sub

Sub test_InlineSection_WriteAndStyleContent()
    Dim inSec As InlineSection
    Set inSec = New InlineSection
    
    With inSec
        .Text = "So strong and emphasised!"
        With .CharacterStyles
            .Push "Normal"
            .Push "Emphasis"
            .Push "Strong"
        End With
        .ISection_WriteContent
        .ISection_StyleContent
    End With
End Sub

Sub test_InterfaceRaisesError()
    Dim sect As New ISection
    sect.WriteContent
End Sub

Sub test_Exception()
    Throw = Errs.NotImplementedException
End Sub

Sub test_FileReader()
    Dim fn As String
    fn = ThisDocument.Path & "\README.md"
    
    Dim fr As IIo
    Set fr = New FileReader
    
    fr.SetReference fn
    Debug.Print "next", fr.NextLine
    Debug.Print "next", fr.NextLine
    Debug.Print "peek", fr.PeekNextLine
    Debug.Print "next", fr.NextLine
    
    fr.SetReference fn
    Debug.Print "peek", fr.PeekNextLine
    Debug.Print "peek", fr.PeekNextLine
    Debug.Print "next", fr.NextLine
    
'    fr.CloseFile
    
End Sub

Sub test_FileReader_Eof()
    Dim fn As String
    fn = ThisDocument.Path & "\README.md"
    
    Dim fr As IIo
    Set fr = New FileReader
    fr.SetReference fn
    
    Do While Not fr.EOF
        fr.NextLine
    Loop
    
    Debug.Print "File read completely"
    
    fr.NextLine
    
End Sub

Sub test_ClassMutability()
    Dim x As IBlockContainer
    Set x = New BlockContainerList
    x.Children.Push New BlockContainerQuote
    
    Debug.Print TypeName(x), "Children: " & x.Children.Count
    Set x = CBlock(x, New BlockLeafBlankLine)
    
    Debug.Print TypeName(x), "Children: " & x.Children.Count
End Sub

Function CBlockContainer(castObject As IBlockContainer, asObject As IBlockContainer) As IBlockContainer
'   Testing cast of one container type to another.
    With castObject.Children
        Do While .Count > 0
            asObject.Children.Push .Pop
        Loop
    End With
    
    Set CBlock = asObject
End Function

Sub test_ListStyle()
    Dim i As Long
    Dim x As New List
    
    Debug.Print "Standard: " & x.IsStandardStyle
    
    For i = 1 To 3
        Debug.Print "Pushing: " & i
        x.Push i
    Next i
    
    Do While x.Count > 0
        Debug.Print "Popping: " & x.Pop
    Loop
    
    Debug.Print vbNewLine & "-----" & vbNewLine
    Debug.Print "Standard: " & x.IsStandardStyle
    
        For i = 1 To 3
        Debug.Print "Pushing: " & i
        x.Push i
    Next i
    
    Debug.Print vbNewLine & "Reversing..."
    x.SetTapeStyle
    Debug.Print "Standard: " & x.IsStandardStyle & vbNewLine
    
    Do While x.Count > 0
        Debug.Print "Popping: " & x.Pop
    Loop
    
End Sub
