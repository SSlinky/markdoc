VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IoFileReader"
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

' Private Backing Store
'-------------------------------------------------------------------------------
Private mFile As Integer
Private mFileIsOpen As Boolean
Private mNextLine As String
Private mFileName As String


' Properties
'-------------------------------------------------------------------------------
Public Property Get IsEoF() As Boolean
'   Returns True if the entire contents has been read.
    IsEoF = EOF(mFile)
End Property


' Methods
'-------------------------------------------------------------------------------
Public Function PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    If mNextLine = "" Then
        mNextLine = ReadNextLine
    End If
    PeekNextLine = mNextLine
End Function

Public Function ReadNextLine() As String
'   Returns the next line to be read and advances the pointer.
    If Not mNextLine = "" Then
        ReadNextLine = mNextLine
        mNextLine = ""
    Else
        ReadNextLine = IoReadNextLine
    End If
End Function

Public Sub OpenStream(fileName As String)
'   Interface implementation to open a file
    If mFileIsOpen Then CloseStream

    mFileName = fileName
    mFile = FreeFile

    Logger.Log "Opening (" & mFile & "): " & mFileName
    If IsEoF Then Throw = Errs.WarnEmptyFile

    Open mFileName For Input Access Read As #mFile
    mFileIsOpen = True
End Sub

Public Sub CloseStream()
'   Closes the file if it is open.
    If Not mFileIsOpen Then Exit Sub
    Logger.Log "Closing (" & mFile & "): " & mFileName

    Close #mFile
    mFileIsOpen = False
End Sub


' Helpers
'-------------------------------------------------------------------------------
Private Function IoReadNextLine() As String
'   Reads the next line in the file.
    If IsEoF Then Throw = Errs.FileReaderEOF
    Line Input #mFile, IoReadNextLine
End Function


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Get IIo_EOF() As Boolean
    IIo_EOF = IsEoF
End Property

Public Sub IIo_OpenStream(fileName As String)
    OpenStream fileName
End Sub

Public Function IIo_PeekNextLine() As String
    IIo_PeekNextLine = PeekNextLine
End Function

Public Function IIo_ReadNextLine() As String
    IIo_ReadNextLine = ReadNextLine
End Function
