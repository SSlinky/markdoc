VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInlineContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Interface")
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IInlineContent
'   Container to hold text, style, and document location information.
'-------------------------------------------------------------------------------

Public Property Let Text(var As String)
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Text() As String
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Set Range(var As Range)
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Range() As Range
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Set Style(var As Style)
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get Style() As Style
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Let SortIndex(var As Long)
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Property Get SortIndex() As Long
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Property

Public Sub AppendText(str As String)
    Throw = Errs.WarnFeatureDeprecated
    Throw = Errs.InterfaceUsedAsObject
End Sub
