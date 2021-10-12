VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Interface")
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IBlock
'   Basic block type that all container and leaf blocks implement.
'-------------------------------------------------------------------------------


Public Property Let IndentationLevel(var As Long)
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get IndentationLevel() As Long
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Sub WriteContent(toDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub StyleContent(styleDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub
