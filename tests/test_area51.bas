Attribute VB_Name = "test_area51"
Option Explicit

Sub fhawl()
    Dim o As New List
    Dim r As Range
    Dim x As Range
    
    With ActiveDocument
        Set r = .Range(0)
        Set x = .Range(Len(.Content) - 1)
        r = "content"
        x = "xContentx"
    End With
    
    o.Push "abc"
    o.Push r
    
    Set r = Nothing
    
    o.Items.Remove 1
    Set r = o.Pop
    Debug.Print r.Text
End Sub

Sub test_InlineSection_WriteAndStyleContent()
    Dim inSec As InlineSection
    Set inSec = New InlineSection
    
    With inSec
        .Text = "So strong and emphasised!"
        With .CharacterStyles
            .Push "Normal"
            .Push "Emphasis"
            .Push "Strong"
        End With
        .ISection_WriteContent
        .ISection_StyleContent
    End With
End Sub

Sub test_InterfaceRaisesError()
    Dim sect As New ISection
    sect.WriteContent
End Sub

Sub test_Exception()
    Throw = Errs.NotImplementedException
End Sub
