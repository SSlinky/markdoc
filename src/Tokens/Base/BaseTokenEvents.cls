VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseTokenEvents"
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
'   Class: BaseTokenEvents
'   An adaptor for event handling in the Base Token.
'-------------------------------------------------------------------------------

' Events
'-------------------------------------------------------------------------------
Public Event OnInvalidate()
Public Event OnSetOpenCharPos(pos As Long)
Public Event OnSetCloseCharPos(pos As Long)


' Private Backing Store
'-------------------------------------------------------------------------------
' Private mToken As ILexerToken


' Constructor
'-------------------------------------------------------------------------------
' Public Sub Initialize(tok As ILexerToken)
' Attribute Initialize.VB_Description = "Acts as a faux constructor. Must be called immediately after initialization."
' '   Acts as a faux constructor. Must be called immediately after initialization.
' '
' '   Args:
' '       tok: The token for event subscription.

'     If Not mToken Is Nothing Then Exit Sub
'     Set mToken = tok
' End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub RaiseOnInvalidate()
    RaiseEvent OnInvalidate
End Sub
Public Sub RaiseOnSetOpenCharPos(pos As Long)
    RaiseEvent OnSetOpenCharPos(pos)
End Sub
Public Sub RaiseOnSetCloseCharPos(pos As Long)
    RaiseEvent OnSetCloseCharPos(pos)
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
' Public Sub ILexerTokenEvents_OnInvalidate()
'     RaiseEvent OnValidate()
' End Sub
' Public Sub ILexerTokenEvents_OnSetOpenCharPos(pos As Long)
'     RaiseEvent OnSetOpenCharPos(pos As Long)
' End Sub
' Public Sub ILexerTokenEvents_OnSetCloseCharPos(pos As Long)
'     RaiseEvent OnSetCloseCharPos(pos As Long)
' End Sub
