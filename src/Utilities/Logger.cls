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
    Debugging
    Information
    Warning
    Exception
End Enum


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
            LevelLookup = ">>BAD LOOKUP>>: "
    End Select
End Function

Public Sub Log(logValue As String, Optional logLevel As Level)
    Debug.Print LevelLookup(logLevel) & logValue
End Sub
