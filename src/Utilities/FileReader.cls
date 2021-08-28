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
Private mFileName As String

Public Function IIo_NextLine() As String
'   Returns the next line to be read and advances the pointer.
    If Not mNextLine = "" Then
        IIo_NextLine = mNextLine
        mNextLine = ""
    Else
        IIo_NextLine = ReadNextLine
    End If
End Function

Public Function IIo_PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    If mNextLine = "" Then
        mNextLine = ReadNextLine
    End If
    IIo_PeekNextLine = mNextLine
End Function

Private Function ReadNextLine() As String
'   Reads the next line in the file.
    If IIo_EOF Then Throw = Errs.FileReaderEOF
    Dim s As String
    Line Input #mFile, s
    ReadNextLine = s
End Function

Public Function IIo_EOF() As Boolean
'   Returns True if the entire contents has been read.
    IIo_EOF = EOF(mFile)
End Function

Public Sub IIo_SetReference(fileName As String)
'   Interface implementation to open a file
    OpenFile fileName
End Sub

Public Sub OpenFile(fileName As String)
'   Opens a file ReadOnly for reading.
    If mFileIsOpen Then CloseFile
    
    mFileName = fileName
    mFile = FreeFile
    
    Logger.Log "Opening (" & mFile & "): " & mFileName
    
    Open mFileName For Input Access Read As #mFile
    mFileIsOpen = True
End Sub

Public Sub CloseFile()
'   Closes the file if it is open.
    If Not mFileIsOpen Then Exit Sub
    Logger.Log "Closing (" & mFile & "): " & mFileName
    
    Close #mFile
    mFileIsOpen = False
End Sub

Private Sub Class_Terminate()
'   Ensure file is closed before terminating.
'   Not even sure this is required, tbh.
    CloseFile
End Sub
