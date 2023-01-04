VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILexerToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: ILexerToken
'   An interface that represents a lexer token.
'-------------------------------------------------------------------------------

' Properties
'-------------------------------------------------------------------------------
Public Property Get Base() As BaseToken
Attribute Base.VB_Description = "The token base."
End Property


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Debug.Print "This should throw an exception!"
End Sub


' Methods
'-------------------------------------------------------------------------------
Public Sub Initialize(lexer As LexerMarkdown)
Attribute Initialize.VB_Description = "Acts as a faux constructor. Must be called immediately after initialization."
'   Acts as a faux constructor. Must be called immediately after initialization.
'
'   Args:
'       lexer: The lexer for event subscription.
End Sub
