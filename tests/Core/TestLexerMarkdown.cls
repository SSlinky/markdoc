VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestLexerMarkdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: TestLexerMarkdown
'   Tests the LexerMarkdown class.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private Const RAWMARKDOWN As String = "Hello, World!"

Private WithEvents mLexerMarkdown As LexerMarkdown
Private mCharsRead As String
Private mCountOnReadCharRaised As Long
Private mCountEOFRaised As Long
Private mTestsPassed As Long
Private mTestsFailed As Long
Private mTokenized As Boolean


' Properties
'-------------------------------------------------------------------------------
Public Property Get TestsPassed() As Long
Attribute TestsPassed.VB_Description = "Number of tests passed."
    TestsPassed = mTestsPassed
End Property

Public Property Get TestsFailed() As Long
Attribute TestsFailed.VB_Description = "Number of tests failed."
    TestsFailed = mTestsFailed
End Property

Public Property Get TestsTotal() As Long
Attribute TestsTotal.VB_Description = "Total number of tests run."
    TestsTotal = mTestsPassed + mTestsFailed
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    set mLexerMarkdown = New LexerMarkdown
End Sub

Private Sub Class_Terminate()
    PrintResults
End Sub

' Methods
'-------------------------------------------------------------------------------
Public Sub PrintResults()
Attribute PrintResults.VB_Description = "Prints the test results."
'   Prints the test results.

    Debug.Print "--- " & mTestsPassed & " passed, " _
                & mTestsFailed & " failed. " _
                & TestsTotal & " total. ---"
End Sub

Public Sub RunAllTests()
Attribute RunAllTests.VB_Description = "Runs all the tests."
'   Runs all the tests.
'
    TestOnReadCharRaisedWhenTokenizing
    TestStringTokenisedMatches
    TestOnEOFRaisedWhenTokenizing
End Sub

Public Sub TestOnReadCharRaisedWhenTokenizing()
Attribute TestOnReadCharRaised.VB_Description = "Tests the OnReadChar event is raised the right number of times."
'   Tests the OnReadChar event is raised the right number of times.

    If Not mTokenized Then RunTokenization RAWMARKDOWN

    If mCountOnReadCharRaised = Len(RAWMARKDOWN) Then
        mTestsPassed = mTestsPassed + 1
        Debug.Print "   PASS: TestOnReadCharRaisedWhenTokenizing"
    Else
        mTestsFailed = mTestsFailed + 1
        Debug.Print "!! FAIL: TestOnReadCharRaisedWhenTokenizing"
    End If
End Sub

Public Sub TestOnEOFRaisedWhenTokenizing()
Attribute TestOnReadCharRaised.VB_Description = "Tests the OnEOF event is raised the right number of times."
'   Tests the OnReadChar event is raised the right number of times.

    If Not mTokenized Then RunTokenization RAWMARKDOWN

    If mCountEOFRaised = 1 Then
        mTestsPassed = mTestsPassed + 1
        Debug.Print "   PASS: TestOnEOFRaisedWhenTokenizing"
    Else
        mTestsFailed = mTestsFailed + 1
        Debug.Print "!! FAIL: TestOnEOFRaisedWhenTokenizing"
    End If
End Sub

Public Sub TestStringTokenisedMatches()
Attribute TestOnReadCharRaised.VB_Description = "Tests the OnReadChar event is sending the right characters."
'   Tests the OnReadChar event is raised the right number of times.

    If Not mTokenized Then RunTokenization RAWMARKDOWN

    If mCharsRead = RAWMARKDOWN Then
        mTestsPassed = mTestsPassed + 1
        Debug.Print "   PASS: TestStringTokenisedMatches"
    Else
        mTestsFailed = mTestsFailed + 1
        Debug.Print "!! FAIL: TestStringTokenisedMatches"
    End If
End Sub

' Helpers
'-------------------------------------------------------------------------------
Private Sub RunTokenization(var As String)
Attribute RunTokenization.VB_Description = "Runs the tokenization."
'   Runs the tokenization.

'   Set up
    mCharsRead = ""
    mCountOnReadCharRaised = 0
    mCountEOFRaised = 0

'   Tokenize    
    mLexerMarkdown.Tokenize var
    mTokenized = True
End Sub


' Event Handlers
'-------------------------------------------------------------------------------
Private Sub mLexerMarkdown_OnReadChar(char As String, pos As Long)
    mCountOnReadCharRaised = mCountOnReadCharRaised + 1
    mCharsRead = mCharsRead & char
End Sub

Private Sub mLexerMarkdown_OnEOF()
    mCountEOFRaised = mCountEOFRaised + 1
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------

