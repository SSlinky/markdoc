Attribute VB_Name = "test_md"
Option Explicit

Sub RunMarkDoc()
    Dim lexer As LexerMarkdown
    Dim stream As IIo
    
    Set lexer = New LexerMarkdown
    Set stream = New IoFileReader

    stream.OpenStream ActiveDocument.Path & "\tests\README.md"
    lexer.ParseMarkdown stream
End Sub
