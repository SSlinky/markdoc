VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InlineSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ISection

'-------------------------------------------------------------------------------
'   Class: InlineSection
'   Container for inline text and styling information.
'   Currently also writes to document so goes against SOLID and needs rethinking
'-------------------------------------------------------------------------------


Public Text As String

Private mRange As Range
Private mStyles As List

Private Sub Class_Initialize()
    Set mStyles = New List
End Sub

Public Property Get CharacterStyles() As List
    Set CharacterStyles = mStyles
End Property

Public Sub ISection_WriteContent()
    With ActiveDocument
        Set mRange = .Range(Len(.Content) - 1)
    End With
    mRange.Text = Me.Text
End Sub

Public Sub ISection_StyleContent()
    With ActiveDocument
        Do While mStyles.Count > 0
            Select Case .Styles(mStyles.Pop)
                Case Is = "Normal"
                    mRange.Bold = False
                    mRange.Italic = False
                Case Is = "Strong"
                    mRange.Bold = True
                Case Is = "Emphasis"
                    mRange.Italic = True
                Case Is = "CodeInline"
                    mRange.Style = .Styles("CodeInline")
                    mRange.Bold = False
                    mRange.Italic = True
            End Select
        Loop
    End With
End Sub
