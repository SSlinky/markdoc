VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: BaseToken
'   Contains base token functionality that all tokens have.
'-------------------------------------------------------------------------------

' Enums
'-------------------------------------------------------------------------------
Enum TokenStateType
    MODE_INIT
    MODE_OPEN
    MODE_CLOSED
    MODE_INVALID
End Enum


' Events
'-------------------------------------------------------------------------------
Public Event OnReadChar(char As String, pos As Long)
Public Event OnNewLine(pos As Long)
Public Event OnEOF()


' Private Backing Store
'-------------------------------------------------------------------------------
Private WithEvents mLexerMarkdown As LexerMarkdown
Private WithEvents mParentTokenEvents As BaseTokenEvents
Private mParentToken As ILexerToken
Private mTokenState As TokenStateType
Private mOpenCharPos As Long
Private mCloseCharPos As Long
Private mPriorityLevel As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Let TokenState(var As TokenStateType)
Attribute TokenState.VB_Description = "The state of the token."
'   The state of the token.
    mTokenState = var
End Property

Public Property Get TokenState() As TokenStateType
    TokenState = mTokenState
End Property

Public Property Get OpenCharPos() As Long
Attribute OpenCharPos.VB_Description = "The pos of the first character in the stream of raw markdown for this block."
'   The pos of the first character in the stream of raw markdown for this block.
    OpenCharPos = mOpenCharPos
End Property

Public Property Get CloseCharPos() As Long
Attribute CloseCharPos.VB_Description = "The pos of the last character in the stream of raw markdown for this token."
'   The pos of the last character in the stream of raw markdown for this token.
    CloseCharPos = mCloseCharPos
End Property

Public Property Get PriorityLevel() As Long
Attribute PriorityLevel.VB_Description = "The priority level of this token."
'   The priority level of this token.
    PriorityLevel = mPriorityLevel
End Property


' Constructor
'-------------------------------------------------------------------------------
Public Sub Initialize( _
    parentToken As ILexerToken, _
    eventsAdaptor As BaseTokenEvents, _
    lexer As LexerMarkdown, _
    priorityLevel As Long)
Attribute Initialize.VB_Description = "Acts as a faux constructor. Must be called immediately after initialization."
'   Acts as a faux constructor. Must be called immediately after initialization.
'
'   Args:
'       pnt: The parent token.
'       lex: The lexer for event subscription.
'       lvl: The priority level for this token.

    If Not mParentToken Is Nothing Then Exit Sub

    mPriorityLevel = priorityLevel
    Set mParentToken = parentToken
    Set mLexerMarkdown = lexer
    Set mParentTokenEvents = eventsAdaptor
End Sub


' Methods
'-------------------------------------------------------------------------------


' Helpers
'-------------------------------------------------------------------------------
Private Function RequireInitialization() As Boolean
Attribute IsInitialized.VB_Description = "Raises an exception if not initialized."
'   Raises an exception if not initialized. This method used by methods
'   to ensure initialization is the very first action.
    If mLexerMarkdown Is Nothing Then
        Debug.Print "Not initialized - TODO implement exception."
    End If
End Function

Private Sub HandleTokenClosedEvent(tok As ILexerToken)
Attribute HandleTokenClosedEvent.VB_Description = "If this token is lower priority and clashes, invalidate the token."
'   If this token is lower priority and clashes, invalidate the token.
'
'   Args:
'       tok: The token that has closed as valid.

    If tok.
End Sub

Private Sub InvalidateMe()
Attribute InvalidateMe.VB_Description = "Invalidates this token."
'   Invalidates this token.
    mLexerMarkdown.UnregisterToken mParentToken
End Sub


' Event Handlers
'-------------------------------------------------------------------------------
Private Sub mParentTokenEvents_OnInvalidate()
    InvalidateMe
End Sub

Private Sub mParentTokenEvents_OnSetOpenCharPos(pos As Long)
    mOpenCharPos = pos
    TokenState = MODE_OPEN
    mLexerMarkdown.RegisterToken mParentToken
End Sub

Private Sub mParentTokenEvents_OnSetCloseCharPos(pos As Long)
Attribute mParentTokenEvents_OnSetCloseCharPos.VB_Description = "Close the token at the position and announce valid close."
'   Close the token at the position and announce valid close.
    mCloseCharPos = pos
    TokenState = MODE_CLOSED
    mLexerMarkdown.CloseToken mParentToken, mPriorityLevel
End Sub

Private Sub mLexerMarkdown_OnCloseToken(tok As ILexerToken, lvl As Long)
Attribute mLexerMarkdown_OnCloseToken.VB_Description = "Handle the close event for another token and invalidate Me if clash."
'   Handle the close event for another token and invalidate Me if clash.
    If Me.TokenState = MODE_INIT Then Exit Sub
    If lvl > mPriorityLevel Then Exit Sub
    If tok.Base Is Me Then Exit Sub
    With tok.Base
        If mOpenCharPos >= .OpenCharPos And mOpenCharPos <= .CloseCharPos Then
            InvalidateMe
        End If
    End With
End Sub

Private Sub mLexerMarkdown_OnReadChar(char As String, pos As Long)
Attribute mLexerMarkdown_OnReadChar.VB_Description = "Bubbles the event up to the next listener."
'   Bubbles the event up to the next listener.
    RaiseEvent OnReadChar(char, pos)
End Sub

Private Sub mLexerMarkdown_OnNewLine(pos As Long)
Attribute mLexerMarkdown_OnNewLine.VB_Description = "Bubbles the event up to the next listener."
'   Bubbles the event up to the next listener.
    RaiseEvent OnNewLine(pos)
End Sub

Private Sub mLexerMarkdown_OnEOF()
Attribute mLexerMarkdown_OnEOF.VB_Description = "Bubbles the event up to the next listener."
'   Bubbles the event up to the next listener.
    RaiseEvent OnEOF()
End Sub
