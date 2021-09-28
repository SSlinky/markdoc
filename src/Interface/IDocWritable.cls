VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDocWritable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IDocWritable
'   An interface that describes an object that can write and style content to a
'   Word document.
'
'   This interface also describes an Inline Block. Inline blocks will only have
'   these methods so there is no need to implement a separate interface.
'-------------------------------------------------------------------------------


Public Sub WriteContent(toDoc As Document)
    Throw = Errs.InterfaceUsedAsObject
End Sub

Public Sub StyleContent()
    Throw = Errs.InterfaceUsedAsObject
End Sub
