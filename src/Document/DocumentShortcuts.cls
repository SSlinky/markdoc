VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: DocumentShortcuts
'   Simpler document attibute management.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mStyles As Object

' Properties
'-------------------------------------------------------------------------------


' Constructor
'-------------------------------------------------------------------------------
Public Sub Attach(doc As Document)
'   Attaches the class to a document.
'
'   Args:
'       doc: The document to attach.
'
   
End Sub

' Methods
'-------------------------------------------------------------------------------
Public Function GetStyle(localName As String) As Style
'   Gets a style from the dictionary. May also lazy create.
'
'   Args:
'       localName: The name of the style to get.
'
'   Returns:
'       A document style if found or created. Normal if not.
'
    
End Function


' Helpers
'-------------------------------------------------------------------------------
Private Sub LoadStyles(doc As Document)
'   Loads the styles so they can be verified and accessed via dictionary.
'
'   Args:
'       doc: The document to load from.
'
    Dim sty As Style
    For Each sty In doc.Styles
        mStyles.Add sty.NameLocal, Nothing
    Next sty

End Sub

Private Sub CreateStyle(localName As String)
'   Attempts to create a style if it's not found in the dictionary.
'
'   Args:
'       localName: The name of the style to create.
'
    Const CODEBLOCK As String = "CodeBlock"
    Const BG_GREY As Long = 16
    
    Dim sty As Style
    Select Case localName
        Case Is = CODEBLOCK
            Logger.Log "Making style: " & STYNAME, Level.Information
            Set sty = forDocument.Styles.Add( _
                Name:=STYNAME, _
                Type:=wdStyleTypeParagraph)
            With sty
                .NoProofing = True
                .NextParagraphStyle = "Normal"
                .BaseStyle = "Normal"
                .NoSpaceBetweenParagraphsOfSameStyle = True
                .Shading.BackgroundPatternColorIndex = BG_GREY
                With .Font
                    .Size = 10
                    .Bold = False
                    .Italic = False
                    .AllCaps = True
                    .Name = "Consolas"
                    .ColorIndex = wdDarkBlue
                End With
            End With
        Case Else
            Throw = Errs.DocumentWarnNoNamedDefaultStyle
    End Select
End Sub