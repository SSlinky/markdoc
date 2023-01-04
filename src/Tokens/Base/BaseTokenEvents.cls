VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseTokenEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
