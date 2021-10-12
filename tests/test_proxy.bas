Attribute VB_Name = "test_proxy"
'@Folder("markDoc.Tests")
Option Explicit

Const GITHUB As String = "https://raw.githubusercontent.com/SSlinky/markdoc/master/tests/"
Const EXAMPLE_PROXYCREDS As String = ",http://myproxyurl.com:1234,myusername,mypassword"


Sub test_RunMarkDoc_from_http()
    DocumentShortcuts.Attach ThisDocument

    Dim lexer As LexerMarkdown
    Dim stream As IFileReader

    Logger.LoggingLevel = Information
    Throw.ThrowLevel = NoLevel

    Set lexer = New LexerMarkdown
    Set stream = New FileReaderHttp

    stream.OpenStream GITHUB & "test_md.md"
    lexer.ParseMarkdown stream
    Set lexer.AttachedDocument = ThisDocument

    ThisDocument.content.Delete
    lexer.WriteDocument
End Sub

Sub test_RunMarkDoc_from_http_behind_proxy()
'   This test requires a private_test_proxy module with a PROXYCREDS constant
'   in the same format as the above EXAMPLE_PROXYCREDS.
'   Which test you run will depend on whether you are behind a proxy or not.
    DocumentShortcuts.Attach ThisDocument

    Dim lexer As LexerMarkdown
    Dim stream As IFileReader
    Dim secrets As String

    Logger.LoggingLevel = Information
    Throw.ThrowLevel = NoLevel

    Set lexer = New LexerMarkdown
    Set stream = New FileReaderHttp
    
    ' secrets = private_test_proxy.PROXYCREDS

    stream.OpenStream GITHUB & "test_md.md" & secrets
    lexer.ParseMarkdown stream
    Set lexer.AttachedDocument = ThisDocument

    ThisDocument.content.Delete
    lexer.WriteDocument
End Sub
