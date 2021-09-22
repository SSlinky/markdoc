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

Private mFile As Integer
Private mFileIsOpen As Boolean
Private mNextLine As String
Private mFileName As String


' Properties
'-------------------------------------------------------------------------------
Public Property Get EoF() As Boolean
'   Returns True if the entire contents has been read.
    IIo_EOF = EoF(mFile)
End Property


' Methods
'-------------------------------------------------------------------------------
Public Function ReadNextLine() As String
'   Reads the next line in the file.
    If IIo_EOF Then Throw = Errs.FileReaderEOF
    Dim s As String
    Line Input #mFile, s
    ReadNextLine = s
End Function

Public Function PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    If mNextLine = "" Then
        mNextLine = ReadNextLine
    End If
    IIo_PeekNextLine = mNextLine
End Function

Public Function ReadNextLine() As String
'   Returns the next line to be read and advances the pointer.
    If Not mNextLine = "" Then
        IIo_NextLine = mNextLine
        mNextLine = ""
    Else
        IIo_NextLine = ReadNextLine
    End If
End Function

Public Sub Open(fileName As String)
'   Interface implementation to open a file
    OpenFile fileName
End Sub


' Interface Implementation
'-------------------------------------------------------------------------------
Public Property Get IIo_EoF() As Boolean
    IIo_EoF = EoF
End Property

Public Sub IIo_Open(fileName As String)
    Call Open(fileName)
End Sub

Public Function IIo_PeekNextLine() As String
    IIo_PeekNextLine = PeekNextLine
End Function

Public Function IIo_ReadNextLine() As String
    IIo_ReadNextLine = ReadNextLine
End Function
