VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IIo
'   An interface that describes base input output methods.
'-------------------------------------------------------------------------------


Public Function NextLine() As String
'   Returns the next line to be read and advances the pointer.
    Throw = Errs.InterfaceUsedAsObject
End Function

Public Function PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    Throw = Errs.InterfaceUsedAsObject
End Function

Public Function EOF() As Boolean
'   Returns True if the entire contents has been read.
    Throw = Errs.InterfaceUsedAsObject
End Function
