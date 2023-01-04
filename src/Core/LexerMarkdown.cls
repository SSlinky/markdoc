VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LexerMarkdown"
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

'-------------------------------------------------------------------------------
'   Class: LexerMarkdown
'   Event driven token generator.
'-------------------------------------------------------------------------------


' Enums
'-------------------------------------------------------------------------------


' Events
'-------------------------------------------------------------------------------
Public Event OnCloseToken(tok As ILexerToken, priorityLevel As Long)
Public Event OnReadChar(char As String, pos As Long)
Public Event OnNewLine(pos As Long)
Public Event OnEOF(pos As Long)


' Private Backing Store
'-------------------------------------------------------------------------------
Private Const LF As Long = 10
Private Const CR As Long = 13
Private mTokens As Collection


' Properties
'-------------------------------------------------------------------------------


' Constructor
'-------------------------------------------------------------------------------
Public Property Set Tokens(var As Collection)
Attribute Tokens.VB_Description = "The tokenised version of the markdown."
    Set mTokens = var
End Property

Public Property Get Tokens() As Collection
    Set Tokens = mTokens
End Property



' Methods
'-------------------------------------------------------------------------------
Public Sub Tokenize(txt As String)
Attribute Tokenise.VB_Description = "Tokenizes Commonmark markdown."
'   Tokenizes Commonmark markdown.
'
'   Args:
'       txt: The input markdown.
'
    InitTokens

    Dim lenTxt As Long
    lenTxt = Len(txt)

    Dim i As Long
    For i = 1 to lenTxt
'       Check for carriage returns or line feeds and raise as a new line event.    
        If Mid(txt, i, 1) = Chr(CR) Then
            If Mid(txt, i + 1, 1) = Chr(LF) Then i = i + 1
            RaiseEvent OnNewLine(i)
        Else
            RaiseEvent OnReadChar(Mid(txt, i, 1), i)
        End If
    Next i

    RaiseEvent OnEOF(i)
End Sub

Public Sub UnregisterToken(token As ILexerToken)
Attribute InvalidateToken.VB_Description = "Stops tracking the token."
'   Stops tracking the token.
'
'   Args:
'       token: The token to stop tracking.
'
'   Raises:
'       Item not found exception if the item isn't found.

    Dim i As Long
    For i = 0 to mTokens.Count - 1
        If mTokens(i) Is token Then
            mTokens.Remove(i)
            Exit For
        End If
    Next i
End Sub

Public Sub RegisterToken(token As ILexerToken)
Attribute RegisterToken.VB_Description = "Registers a new token for tracking."
'   Registers a new token for tracking.
'
'   Args:
'       token: The token to begin tracking.

    mTokens.Add TokenFactory(token)
End Sub

Public Sub CloseToken(tok As ILexerToken, priorityLevel As Long)
Attribute CloseToken.VB_Description = "Announces that a token has closed so that others can respond."
'   Announces that a token has closed so that others can respond.
'   Lower priority tokens that clash should be invalidated.
'
'   Args:
'       tok: The token that has closed as valid.

    RaiseEvent OnCloseToken(tok, priorityLevel)
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub InitTokens()
Attribute InitTokens.VB_Description = "Initialises the tokens."
'   Initialises the tokens and registers one of each kind.
'
    Set mTokens = New Collection

    RegisterToken "InlineCodeToken"
End Sub

Private Function TokenFactory(tok As Variant) As ILexerToken
Attribute TokenFactory.VB_Description = "Sets up a new token and returns it."
'   Sets up a new token and returns it.
'
'   Args:
'       tok: A token object or its type as a string.
'
'   Returns:
'       An initialized token.
'
'   Raises:
'       Token not registered exception if the token could not be found.

    Dim tType As String
    tType = TypeName(tok)
    If tType = "String" Then tType = tok

    Dim t As ILexerToken
    Select Case tType
        Case Is = "InlineCodeToken":
            Set t = New InlineCodeToken
        Case Else:
            ' TODO Raise exception
    End Select
    t.Initialize Me
    Set TokenFactory = t
End Function


' Event Handlers
'-------------------------------------------------------------------------------
