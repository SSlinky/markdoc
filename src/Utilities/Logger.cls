VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Static Class: Logger
'   Allows for better logging later.
'-------------------------------------------------------------------------------

Public Enum Level
    NoLevel
    Debugging
    Information
    Warning
    Exception
    NeverThrow
End Enum

' Private Backing Store
'-------------------------------------------------------------------------------
Private mLoggingLevel As Level


' Properties
'-------------------------------------------------------------------------------
Public Property Let LoggingLevel(var As Level)
    Logger.Log "Logging set to " & Mid(LevelLookup(var), 1, 4), NoLevel
    mLoggingLevel = var
End Property

Public Property Get LoggingLevel() As Level
    LoggingLevel = mLoggingLevel
End Property


' Methods
'-------------------------------------------------------------------------------
Private Function LevelLookup(logLevel As Level) As String
    Select Case logLevel
        Case Is = Level.Information
            LevelLookup = "INFO: "
        Case Is = Level.Debugging
            LevelLookup = "DEBG: "
        Case Is = Level.Warning
            LevelLookup = "WARN: "
        Case Is = Level.Exception
            LevelLookup = "EXPT: "
        Case Else
            LevelLookup = "~~~~~"
    End Select
End Function

Public Sub Log(logValue As String, Optional logLevel As Level)
    If logLevel < LoggingLevel Then Exit Sub
    Debug.Print LevelLookup(logLevel) & logValue
End Sub
