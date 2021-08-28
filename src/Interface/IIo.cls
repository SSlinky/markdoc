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
    Throw = Errs.InterfaceUsedAsObject
End Function
