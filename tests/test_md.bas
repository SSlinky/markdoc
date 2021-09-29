Sub RunMarkDoc()
    Dim lexer As LexerMarkdown
    Dim stream As IIo
    
    Set lexer = New LexerMarkdown
    Set stream = New FileReader

    stream.Open(ActiveDocument.Path & "\tests\README.md")
    lexer.ParseMarkdown(stream)

    
End Sub
