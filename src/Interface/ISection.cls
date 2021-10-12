VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Interface")
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: ISection
'   An interface that describes base Section methods
'-------------------------------------------------------------------------------


Public Sub WriteContent(toDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub StyleContent(styleDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub
