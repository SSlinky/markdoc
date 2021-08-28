VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: ISection
'   An interface that describes base Section methods
'-------------------------------------------------------------------------------


Public Sub WriteContent()
    Err.Raise 900, "markDoc:ISection", "Interface not implemented."
End Sub

Public Sub StyleContent()
    Err.Raise 900, "markDoc:ISection", "Interface not implemented."
End Sub
