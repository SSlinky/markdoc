VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILexerTokenEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: ILexerTokenEvents
'   An interface that describes the events raised by a Lexer Token
'-------------------------------------------------------------------------------

' Enums
'-------------------------------------------------------------------------------


' Events
'-------------------------------------------------------------------------------
Public Sub OnInvalidate(): End Sub
Public Sub OnSetOpenCharPos(pos As Long): End Sub
Public Sub OnSetCloseCharPos(pos As Long): End Sub


' Private Backing Store
'-------------------------------------------------------------------------------


' Properties
'-------------------------------------------------------------------------------


' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------


' Helpers
'-------------------------------------------------------------------------------


' Event Handlers
'-------------------------------------------------------------------------------


' Interface Implementation
'-------------------------------------------------------------------------------

