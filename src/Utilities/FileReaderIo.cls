VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReaderIo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Utilities")
Option Explicit
Implements IFileReader

'-------------------------------------------------------------------------------
'   Class: FileReaderIo
'   Reads the contents of a file on the system line by line.
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


' Constructor
'-------------------------------------------------------------------------------
Private Sub Class_Terminate()
    CloseStream
End Sub


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

    Open mFileName For Input Access Read As #mFile
    Logger.Log "Opened (" & mFile & "): " & mFileName, Level.Information
    If IsEoF Then Throw = Errs.FileReaderWarnEmptyFile

    mFileIsOpen = True
End Sub

Public Sub CloseStream()
'   Closes the file if it is open.
    If Not mFileIsOpen Then Exit Sub
    Logger.Log "Closing (" & mFile & "): " & mFileName, Level.Information

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
Public Property Get IFileReader_EOF() As Boolean
    IFileReader_EOF = IsEoF
End Property

Public Sub IFileReader_OpenStream(fileName As String)
    OpenStream fileName
End Sub

Public Function IFileReader_PeekNextLine() As String
    IFileReader_PeekNextLine = PeekNextLine
End Function

Public Function IFileReader_ReadNextLine() As String
    IFileReader_ReadNextLine = ReadNextLine
End Function
