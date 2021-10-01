VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Throw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Static Class: Throw
'   Adds a modern language feel to throwing errors.
'-------------------------------------------------------------------------------

Public Enum Errs
'   Error types enum. Each error should
'   have an equivalent description in ErrsLookup.

'   Application
    InterfaceUsedAsObject
    NotImplementedException
    WarnNotImplementedException

'   Block
    BlockNotCreated
    BlockTypeMismatch

'   File Reader
    FileReaderEOF
    FileReaderWarnEmptyFile

'   Lexer
    LexerWarnNothingToWrite    

'   Parser
    ParserIndentLevelOutsideRange
    
'   Utility
    UtilsBadArguments
    UtilsStringIndentNotBlank
    UtilsWarnDestructiveCast
End Enum


' Private Backing Store
'-------------------------------------------------------------------------------
Private mThrowLevel As Level


' Properties
'-------------------------------------------------------------------------------
Public Property Let Exception(errType As Errs)
    Attribute Exception.VB_UserMemId = 0
    Dim errDesc As String
    Dim logLevel As Level

    errDesc = ErrsLookup(errType)
    logLevel = ErrsLevelLookup(errType)
    
    Logger.Log errDesc, logLevel
    If logLevel >= ThrowLevel And logLevel > Level.Information Then
        Err.Raise vbObjectError + errType + 1, Description:=errDesc
    End If
End Property

Public Property Let ThrowLevel(var As Level)
    Logger.Log "ThrowLevel set to " & Mid(LevelLookup(var), 1, 4), NoLevel
    mThrowLevel = var
End Property

Public Property Get ThrowLevel() As Level
    ThrowLevel = mThrowLevel
End Property


' Helpers
'-------------------------------------------------------------------------------
Private Function ErrsLookup(errType As Errs) As String
    Select Case errType
        Case Is = Errs.InterfaceUsedAsObject
            ErrsLookup = "Interface cannot be concrete type."
        Case Is = Errs.WarnNotImplementedException
            ErrsLookup = "Not yet implemented."
        Case Is = Errs.NotImplementedException
            ErrsLookup = "Not yet implemented."
        Case Is = Errs.BlockNotCreated
            ErrsLookup = "Block must be created using Create method."
        Case Is = Errs.BlockTypeMismatch
            ErrsLookup = "Block type mismatch."
        Case Is = Errs.FileReaderEOF
            ErrsLookup = "End of file."
        Case Is = Errs.FileReaderWarnEmptyFile
            ErrsLookup = "File is empty."
        Case Is = Errs.LexerWarnNothingToWrite
            ErrsLookup = "Cannot write. No markdown parsed."
        Case Is = Errs.ParserIndentLevelOutsideRange
            ErrsLookup = "Cannot set new indentation level."
        Case Is = Errs.UtilsBadArguments
            ErrsLookup = "Arguments passed not valid."
        Case Is = Errs.UtilsStringIndentNotBlank
            ErrsLookup = "String indentation must be white space."
        Case Is = Errs.UtilsWarnDestructiveCast
            ErrsLookup = "Cast resulted in loss of information."
        Case Else
            ErrsLookup = "Internal error " & errType & "."
    End Select
End Function

Private Function ErrsLevelLookup(errType As Errs) As Level
'   Lookup to get error severity.
'
'   Args:
'       errType: The type of exception being thrown.
'
'   Returns:
'       A logger level.
'
    Select Case errType
        Case Is = Errs.WarnNotImplementedException, _
                  Errs.FileReaderWarnEmptyFile, _
                  Errs.LexerWarnNothingToWrite, _
                  Errs.UtilsWarnDestructiveCast
            ErrsLevelLookup = Level.Warning
        Case Else
            ErrsLevelLookup = Level.Exception
    End Select
End Function
