VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Throw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("markDoc.src.Utilities")
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
'   Static Class: Throw
'   Adds a modern language feel to throwing errors.
'-------------------------------------------------------------------------------

Public Enum Errs
'   Error types enum. Each error should
'   have an equivalent description in ErrsLookup.

'   Application
    ApplicationValueError
    InterfaceUsedAsObject
    NotImplementedException
    WarnNoNamedDefaultStyle
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
'   Throw exception if the error level is high enough.
'   Never throw for Information or NoLevel.
'   Always throw for Critical.
    If logLevel >= ThrowLevel Or logLevel = Level.Critical Then
        If logLevel > Level.Information And logLevel <> Level.NoLevel Then
            Err.Raise vbObjectError + errType + 1, Description:=errDesc
        End If
    End If
End Property

Public Property Let ThrowLevel(var As Level)
    Dim loglev As String
    loglev = Mid(Logger.LevelLookup(var), 1, 4)
    Logger.Log "ThrowLevel set to " & loglev, NoLevel
    mThrowLevel = var
End Property

Public Property Get ThrowLevel() As Level
    ThrowLevel = mThrowLevel
End Property


' Helpers
'-------------------------------------------------------------------------------
Private Function ErrsLookup(errType As Errs) As String
    Select Case errType
        Case Is = Errs.ApplicationValueError
            ErrsLookup = "Value error."
        Case Is = Errs.InterfaceUsedAsObject
            ErrsLookup = "Interface cannot be concrete type."
        Case Is = Errs.NotImplementedException
            ErrsLookup = "Not yet implemented."
        Case Is = Errs.WarnNoNamedDefaultStyle
            ErrsLookup = "Could not create default style."
        Case Is = Errs.WarnNotImplementedException
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
        Case Is = Errs.WarnNoNamedDefaultStyle, _
                  Errs.WarnNotImplementedException, _
                  Errs.FileReaderWarnEmptyFile, _
                  Errs.LexerWarnNothingToWrite, _
                  Errs.UtilsWarnDestructiveCast
            ErrsLevelLookup = Level.Warning
        Case Is = Errs.ApplicationValueError
            ErrsLevelLookup = Level.Critical
        Case Else
            ErrsLevelLookup = Level.Exception
    End Select
End Function
