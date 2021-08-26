VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'   Class: BlockSection
'   Container for block (paragraph) text and styling information.
'   A block will contain at least one InlineSection. This container is for
'   logically separating out paragraphs in the document.
'-------------------------------------------------------------------------------

Private mInsecs As List

Private Sub Class_Initialize()
    Set mInsecs = New List
End Sub

Public Sub WriteText()
    Dim n As Long
    n = mInsecs.Count
    
    Dim inSec As InlineSection
    
    Dim i As Long
    For i = 1 To n
        Set inSec = mInsecs.Pop
        inSec.WriteText
        mInsecs.Push inSec
    Next i
End Sub
