VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileReaderHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("markDoc.src.Utilities")
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

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
Attribute PeekNextLine.VB_Description = "Returns the next line to be read without advancing the pointer."
    If IsEoF Then Throw = Errs.FileReaderEOF
    PeekNextLine = mLines.Peek
End Function

Public Function ReadNextLine() As String
'   Returns the next line to be read and advances the pointer.
    If IsEoF Then Throw = Errs.FileReaderEOF
    ReadNextLine = mLines.Pop
End Function

Public Sub OpenStream(args As String)
'   Loads file into backing store so that it can be read line by line.
'   The args parameter is split and read. Not ideal but needs to be generic.
'
'   Args:
'       args: (excl proxy) The url location of the file.
'             (incl proxy) url,proxy url,username,password.
'
    Const HTTPREQUEST_PROXYSETTING_PROXY As Long = 2

    Dim resource As String
    Dim proxyUrl As String
    Dim proxyUsr As String
    Dim proxyPwd As String
    Dim spltArgs() As String
    Dim useProxy As Boolean

'   Set up the GET request parameters.
    spltArgs = Split(args, ",")
    If UBound(spltArgs) = 3 Then
        useProxy = True
        proxyUrl = spltArgs(1)
        proxyUsr = spltArgs(2)
        proxyPwd = spltArgs(3)
    ElseIf UBound(spltArgs) <> 0 Then
        Throw = Errs.UtilsBadArguments
    End If
    resource = spltArgs(0)

    Dim httpResult As String
    Set mLines = New List

'   Send the GET request.
    With CreateObject("MSXML2.ServerXMLHTTP")
        .Open "GET", resource, False
        If useProxy Then
            .SetProxy HTTPREQUEST_PROXYSETTING_PROXY, proxyUrl, ""
            .SetProxyCredentials proxyUsr, proxyPwd
        End If
        .Send
'       Convert the result from an ANSI byte array to unicode.
        httpResult = StrConv(.ResponseBody, vbUnicode)
    End With

    Dim fileLines() As String
    fileLines = Split(httpResult, vbLf)

'   Add the lines to the backing store so they can be iterated through.
    Dim i As Long
    For i = 0 To UBound(fileLines)
        mLines.Push fileLines(i)
    Next i
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
