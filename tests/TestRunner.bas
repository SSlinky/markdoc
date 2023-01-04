Attribute VB_Name = "TestRunner"
Option Explicit

Public Sub RunTestLexerMarkdown()
Attribute RunTestLexerMarkdown.VB_Description = "Runs all Lexer tests."
'   Runs all Lexer tests.
    Dim runner As New TestLexerMarkdown
    runner.RunAllTests
End Sub
