VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReaderHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IFileReader

'-------------------------------------------------------------------------------
'   Class: FileReaderHttp
'   Reads the contents of a file at an http location.
'-------------------------------------------------------------------------------

' Private Backing Store
'-------------------------------------------------------------------------------
Private mLines As List
Private mNextLine As String


' Properties
'-------------------------------------------------------------------------------
Public Property Get IsEoF() As Boolean
'   Returns True if the entire contents has been read.
    IsEoF = mLines.Count = 0
End Property


' Methods
'-------------------------------------------------------------------------------
Public Function PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    If IsEoF Then Throw = Errs.FileReaderEOF
    PeekNextLine = mLines.Peek
End Function

Public Function ReadNextLine() As String
'   Returns the next line to be read and advances the pointer.
    If IsEoF Then Throw = Errs.FileReaderEOF
    ReadNextLine = mLines.Pop
End Function

Public Sub OpenStream(httpPath As String)
'   Loads file into backing store so that it can be read line by line.
    Throw = Errs.NotImplementedException
End Sub

Public Sub CloseStream()
'   Discards the file.
    Set mLines = New List
End Sub


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
