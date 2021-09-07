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
    NotImplementedException
    InterfaceUsedAsObject
    UtilsBadArguments
End Enum

Private Function ErrsLookup(errType As Errs) As String
    Select Case errType
        Case Is = Errs.NotImplementedException
            ErrsLookup = "Not yet implemented."
        Case Is = Errs.InterfaceUsedAsObject
            ErrsLookup = "Interface cannot be concrete type."
        Case Is = Errs.UtilsBadArguments
            ErrsLookup = "Arguments passed not valid."
        Case Else
            ErrsLookup = "Internal error " & errType & "."
    End Select
End Function

Public Property Let Exception(errType As Errs)
    Dim errDesc As String
    errDesc = ErrsLookup(errType)
    
    Logger.Log errDesc, Level.Exception
    Err.Raise vbObjectError + errType + 1, Description:=errDesc
End Property
