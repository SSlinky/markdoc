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


Public Function CountSubstr(str As String, substr As String) As Long
    CountSubstr = (Len(str) - Len(Replace(str, substr, ""))) / Len(substr)
End Function

Public Function GetIndentationLevel(line As String) As Long
    If Trim(line) = "" Then
        GetIndentationLevel = 0
        Exit Function
    End If

    Dim c As String * 1
    Dim i As Long
    
    For i = 1 to Len(line)
        If Not Mid(line, i, 1) = " " Then
            GetIndentationLevel = i / 4
            Exit Function
        End If
    Next i
End Function

Function SuperTrim(str As String, Optional TrimChars As String) As String
'   Removes any character in TrimChars from the start or end of the line.
    If Timchars = "" Then
        TrimChars = vbNewLine _ ' Carriage return and line feed
                    & Chr(160)  ' nbsp
    End If
    
End Function

Function CleanString(str As String, Optional dirtyChars As String, Optional cleanChars As String) As String
'   Converts any dirtyChar into a cleanChar.
'   You cannot pass dirtyChars without cleanChars and they must be the same length.
    If Len(dirtyChars) = Len(cleanChars) Then Throw = UtilsBadArguments
End Function