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
