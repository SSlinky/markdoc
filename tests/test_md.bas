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
    Dim lexer As LexerMarkdown
    Dim stream As IIo
    
    Set lexer = New LexerMarkdown
    Set stream = New IoFileReader

    stream.OpenStream ActiveDocument.Path & "\tests\test_empty.md"
End Sub
