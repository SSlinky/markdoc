VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineCodeToken"
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
'   Class: InlineCodeToken
'   A token that represents a section of inline code.
'-------------------------------------------------------------------------------

' Enums
'-------------------------------------------------------------------------------
Enum InlineCodeTokenStateType
    MODE_INIT       ' Watching for first open character.
    MODE_ESCAPED    ' Next character can't initialize opening.
    MODE_OPENING    ' Watching for valid sequence of open characters.
    MODE_OPEN       ' Watching for first close character.
    MODE_CLOSING    ' Watching for valid sequence of close characters.
    MODE_CLOSED     ' Token is validated.
End Enum


' Events
'-------------------------------------------------------------------------------
' Public Event OnInvalidate()
' Public Event OnSetOpenCharPos(pos As Long)
' Public Event OnSetCloseCharPos(pos As Long)


' Private Backing Store
'-------------------------------------------------------------------------------
Private WithEvents mBase As BaseToken
Private mEvents As BaseTokenEvents
Private mTokState As InlineCodeTokenStateType
Private mPosOpenStart As Long
Private mLenOpenFence As Long
Private mLenCloseFence As Long


' Properties
'-------------------------------------------------------------------------------
Public Property Get Base() As BaseToken
Attribute Base.VB_Description = "The token base."
    Set Base = mBase
End Property


' Constructor
'-------------------------------------------------------------------------------


' Methods
'-------------------------------------------------------------------------------
Public Sub Initialize(lexer As LexerMarkdown)
Attribute Initialize.VB_Description = "Acts as a faux constructor. Must be called immediately after initialization."
'   Acts as a faux constructor. Must be called immediately after initialization.
'
'   Args:
'       lexer: The lexer for event subscription.

    If Not mBase Is Nothing Then Exit Sub
    Set mBase = New BaseToken
    Set mEvents = New BaseTokenEvents
    mBase.Initialize parentToken:=Me, _
                     eventsAdaptor:=mEvents, _
                     lexer:=lexer, _
                     priorityLevel:=2
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Sub RequireInitialization()
Attribute RequireInitialization.VB_Description = "Raises an exception if not initialized."
'   Raises an exception if not initialized. This method used by methods
'   to ensure initialization is the very first action.
    If mBase Is Nothing Then
        Debug.Print "Not initialized - TODO implement exception."
    End If
End Sub

Private Sub FeedStateMachine(char As String, pos As Long)
Attribute FeedStateMachine.VB_Description = "Runs the character through the state machine."
'   Runs the character through the state machine.
'
'   Args:
'       char: The character being processed.
'       pos: The position in raw markdown of the character.

    Select Case mTokState
        Case Is = MODE_CLOSED:
'           Do nothing.        
        Case Is = MODE_INIT:
'           Test for an excape character or an open fence character.        
            Select Case char
                Case Is = "\":
                    mTokState = MODE_ESCAPED
                Case Is = "`":
                    mPosOpenStart = pos
                    mTokState = MODE_OPENING
            End Select
        Case Is = MODE_ESCAPED:
'           Toggle escape off regardless of what we get.        
            mTokState = MODE_INIT
        Case Is = MODE_OPENING:
'           Watch for a non-open fence character.        
            If char <> "`" Then OpenToken pos
        Case Is = MODE_OPEN:
'           Watch for a close fence character.        
            If char = "`" Then
                mLenCloseFence = 1
                mTokState = MODE_CLOSING
            End If
        Case Is = MODE_CLOSING:
'           Increment close characters until we get a non-close character.
            If char = "`" Then
                mLenCloseFence = mLenCloseFence + 1
            Else
'               Close as valid if open and close fences match.
                If mLenCloseFence = mLenOpenFence Then
                    mEvents.RaiseOnSetCloseCharPos pos - 1
'               Otherwise reset to open.
                Else
                    mTokState = MODE_OPEN
                End If          
            End If
    End Select
End Sub

Private Sub OpenToken(pos As Long)
Attribute OpenToken.VB_Description = "Opens the token at the position."
'   Opens the token at the position.
'
'   Args:
'       pos: The position in raw markdown to open at.

    mEvents.RaiseOnSetOpenCharPos(mPosOpenStart)
    mLenOpenFence = pos - mPosOpenStart
    mTokState = MODE_OPEN
End Sub

Private Sub ValidateEndOfLine(pos As Long)
Attribute ValidateEndOfLine.VB_Description = "Closes or invalidates the token."
'   Closes or invalidates the token.
    Select Case True
'       Close as valid token.    
        Case Is = mTokState = MODE_CLOSING And mLenCloseFence = mLenOpenFence:
            mEvents.RaiseOnSetCloseCharPos pos
'       Invalidate and unregister the token.            
        Case Is = mBase.TokenState <> TokenStateType.MODE_CLOSED:
            mEvents.RaiseOnInvalidate
    End Select
End Sub


' Event Handlers
'-------------------------------------------------------------------------------
Private Sub mBase_OnReadChar(char As String, pos As Long)
    FeedStateMachine char, pos
End Sub

Private Sub mBase_OnNewLine(pos As Long)
    ValidateEndOfLine
End Sub

Private Sub mBase_OnEOF(pos As Long)
    ValidateEndOfLine
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Sub ILexerToken_Initialize(lexer As LexerMarkdown)
    Initialize lexer
End Sub

Public Property Get ILexerToken_Base() As BaseToken
    Set ILexerToken_Base = Base()
End Sub
