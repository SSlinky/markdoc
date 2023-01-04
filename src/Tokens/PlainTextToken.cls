VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlainTextToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

Option Explicit
Implements ILexerToken

'-------------------------------------------------------------------------------
'   Class: PlainTextToken
'   A special token for text that doesn't tokenize into anything else.
'   This token is a simple container manually created by the Lexor.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mBase As BaseToken


' Methods
'-------------------------------------------------------------------------------
Public Sub Initialize(lexer As LexerMarkdown, openPos As Long, closePos As Long)
Attribute Initialize.VB_Description = "Acts as a faux constructor. Must be called immediately after initialization."
'   Acts as a faux constructor. Must be called immediately after initialization.
'   A PlainTextToken will always be valid and the open and close positions
'   will be set by the calling Lexer.
'
'   Args:
'       lexer: The lexer for event subscription.

    If Not mBase Is Nothing Then Exit Sub
    Set mBase = New BaseToken
    Set mEvents = New BaseTokenEvents
    mBase.Initialize parentToken:=Me, _
                     eventsAdaptor:=mEvents, _
                     lexer:=lexer, _
                     priorityLevel:=2, _
                     openPos:=openPos, _
                     closePos:=closePos
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub ILexerToken_Initialize(lexer As LexerMarkdown)
'   Do nothing
End Sub

Public Property Get ILexerToken_Base() As BaseToken
    Set ILexerToken_Base = Base()
End Sub
