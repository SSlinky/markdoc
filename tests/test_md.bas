Attribute VB_Name = "test_md"
Option Explicit

Sub test_RunMarkDoc()
    Dim lexer As LexerMarkdown
    Dim stream As IIo
    
    Set lexer = New LexerMarkdown
    Set stream = New IoFileReader

    stream.OpenStream ActiveDocument.Path & "\tests\test_md.md"
    lexer.ParseMarkdown stream
End Sub

Sub test_EmptyDoc()
    Dim stream1 As IIo
    Dim stream2 As IIo
    
    Set stream1 = New IoFileReader
    Set stream2 = New IoFileReader
    
    stream1.OpenStream ActiveDocument.Path & "\tests\test_empty.md"
    stream2.OpenStream ActiveDocument.Path & "\tests\test_md.md"
End Sub
