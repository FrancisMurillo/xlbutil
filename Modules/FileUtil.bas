Attribute VB_Name = "FileUtil"
Public Enum Encoding
    ANSI
    UNICODE
    UNICODEBOM
    UTF8
End Enum


'# Opens the file dialog to find a file and returns the path
'# If it was canceled, it returns the empty string
Public Function BrowseFile() As String
    Dim intChoice As Integer
    Dim strPath As String
    
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'determine what choice the user made
    If intChoice <> 0 Then
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        'print the file path to sheet 1
        BrowseFile = strPath
    Else
        BrowseFile = ""
    End If
End Function

Public Function DoesFileExists(FilePath As String) As Boolean
    DoesFileExists = Dir(FilePath, vbNormal) <> ""
End Function

'# Gets the whole contents of a text file
'! Raises an exception when it cannot be parsed
Public Function GetText(sFile As String) As String
    Dim FileEncoding As Encoding
    FileEncoding = DetectEncoding(sFile)

    Select Case FileEncoding
        Case ANSI, UNICODE
            GetText = ReadFile(sFile)
        Case Else
            GetText = ReadFileAsUTF8(sFile)
    End Select
End Function

Public Function ReadFile(sFile As String) As String
    Dim nSourceFile As Integer, sText As String

   ''Close any open text files
   Close

   ''Get the number of the next free text file
   nSourceFile = FreeFile

   ''Write the entire file to sText
   Open sFile For Input As #nSourceFile
   sText = Input$(LOF(1), 1)
   Close

   ReadFile = sText
End Function

'# Reads a file via UTF
Public Function ReadFileAsUTF8(sFile As String) As String
    Dim objStream, strData
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (sFile)
    ReadFileAsUTF8 = objStream.ReadText()
End Function

'# Determines a file encoding
Public Function DetectEncoding(ByVal FileName As String) As Encoding
  Dim b1 As Byte, b2 As Byte, C As String
  Open FileName For Binary As #1
  Get #1, , b1
  Get #1, , b2
  Close #1
  
  If b1 = &HFF And b2 = &HFE Then C = UNICODE Else _
  If b1 = &HFE And b2 = &HFF Then C = UNICODEBOM Else _
  If b1 = &HEF And b2 = &HBB Then C = UTF Else C = ANSI
  
  DetectEncoding = C
End Function



