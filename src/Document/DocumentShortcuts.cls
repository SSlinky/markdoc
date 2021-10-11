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
Private mDoc As Document
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
    Set mDoc = doc

'   Create and populate the styles dictionary.
    LoadStyles doc
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
    If Not mStyles.Exists(localName) Then
'       Attempt to create the style if it does not exist.
        Logger.Log "Style " & localName & " does not exist.", _
            Level.Information
        CreateStyle localName
'       Return Normal if we couldn't create.
'       Function calls itself just in case Normal has been deleted.
        If Not mStyles.Exists(localName) Then
            Set GetStyle = GetStyle("Normal")
            Exit Function
        End If
    End If
    Set GetStyle = mDoc.Styles(localName)
End Function


' Helpers
'-------------------------------------------------------------------------------
Private Sub LoadStyles(doc As Document)
'   Loads the styles so they can be verified and accessed via dictionary.
'
'   Args:
'       doc: The document to load from.
'
    Set mStyles = CreateObject("Scripting.Dictionary")
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
'   Font constants
    Const FONT_NAME As String = "Consolas"
    Const FONT_SIZE As Long = 10

'   Colour constants    
    Const FG_ORANGE As Long = -721354753
    Const BG_GREY As Long = 16

'   Style name constants
    Const NORMAL As String = "Normal"
    Const CODEBLOCK As String = "CodeBlock"
    Const CODEINLINE As String = "CodeInline"

    Dim sty As Style
    Select Case localName
        Case Is = CODEBLOCK
            Logger.Log "Making style: " & localName, Level.Information
            Set sty = mDoc.Styles.Add( _
                Name:=localName, _
                Type:=wdStyleTypeParagraph)
            With sty
                .NoProofing = True
                .NextParagraphStyle = NORMAL
                .BaseStyle = NORMAL
                .NoSpaceBetweenParagraphsOfSameStyle = True
                .Shading.BackgroundPatternColorIndex = BG_GREY
                With .Font
                    .Size = FONT_SIZE
                    .Bold = False
                    .Italic = False
                    .AllCaps = False
                    .Name = FONTNAME
                    .ColorIndex = wdDarkBlue
                End With
            End With
            mStyles.Add localName, sty

        Case Is = CODEINLINE
            Set sty = forDocument.Styles.Add( _
            Name:=STYNAME, _
            Type:=wdStyleTypeCharacter)
            With sty
                .NoProofing = True
                With .Font
                    .Size = FONT_SIZE
                    .Bold = False
                    .Italic = False
                    .AllCaps = False
                    .Name = FONTNAME
                    .Color = FG_ORANGE
                End With
            End With

        Case Is = NORMAL
            Throw = Errs.NotImplementedException

        Case Else
            Throw = Errs.WarnNoNamedDefaultStyle

    End Select
End Sub