VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Static Class: Utils
'   Contains generic static helper functions.
'-------------------------------------------------------------------------------

Const NBSP As Long = 160

Public Function CountSubstr(str As String, substr As String) As Long
    CountSubstr = (Len(str) - Len(Replace(str, substr, ""))) / Len(substr)
End Function

Public Function GetIndentationLevel(line As String) As Long
    If Trim(line) = "" Then
        GetIndentationLevel = 0
        Exit Function
    End If

    Dim i As Long    
    For i = 1 to Len(line)
        If Not Mid(line, i, 1) = " " Then
            GetIndentationLevel = i / 4
            Exit Function
        End If
    Next i
End Function

Function SuperTrim(str As String, Optional trimChars As String) As String
'   Removes any character in trimChars from the start or end of the line.
    Dim i As Long
    
'   Default trim characters
    If trimChars = "" Then trimChars = vbNewLine & Chr(NBSP) & " "
    
'   Trim left
    For i = 1 To Len(str)
        If Not StrIn(Mid(str, i, 1), trimChars) Then Exit For
    Next i
    If i > 1 Then str = Right(str, Len(str) - i + 1)

'   Trim right
    For i = Len(str) To 1 Step -1
        If Not StrIn(Mid(str, i, 1), trimChars) Then Exit For
    Next i
    If i < Len(str) Then str = Left(str, i)
    
    SuperTrim = str
End Function

Function CleanString(str As String, Optional dirtyChars As String, Optional cleanChars As String) As String
'   Converts any single dirtyChar into a cleanChar.
'   You cannot pass dirtyChars without cleanChars and they must be the same length.
   If Len(dirtyChars) = Len(cleanChars) Then Throw = UtilsBadArguments
    
'   Set up defaults if nothing passed
    If dirtyChars = "" Then
        dirtyChars = Chr(NBSP)
        cleanChars = " "
    End If
    
    Dim i As Long
    For i = 1 To Len(dirtyChars)
        str = Replace(str, Mid(dirtyChars, i, 1), Mid(cleanChars, i, 1))
    Next i
    CleanString = str
End Function

Function StrIn(str As String, subStr As String) As Boolean
'   Helper for Instr for when you don't care where subStr is.
    StrIn = Instr(str, subStr) > 0
End Function