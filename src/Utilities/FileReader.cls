VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IIo

'-------------------------------------------------------------------------------
'   Class: FileReader
'   Reads the contents of a file line by line.
'-------------------------------------------------------------------------------

Private mFile As Integer
Private mFileIsOpen As Boolean
Private mNextLine As String

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

Public Sub OpenFile(fileName As String)
'   Opens a file ReadOnly for reading.
    If mFileIsOpen Then CloseFile
    mFile = FreeFile
    
    Open fileName For Input Access Read As #mFile
    mFileIsOpen = True
End Sub

Public Sub CloseFile()
'   Closes the file if it is open.
    If Not mFileIsOpen Then Exit Sub
    Close #mFile
    mFileIsOpen = False
End Sub

Private Sub Class_Terminate()
'   Ensure file is closed before terminating.
'   Not even sure this is required, tbh.
    CloseFile
End Sub
