Attribute VB_Name = "BookUtil"
'# Gets all the sheet name available
Public Function GetSheetNames(Book As Workbook) As Variant
    Dim Arr_ As Variant, Index As Long
    Arr_ = ArrayUtil.CreateWithSize(Book.Worksheets.Count)
    For Index = 0 To UBound(Arr_)
        Arr_(Index) = Book.Worksheets(Index + 1).Name
    Next

    GetSheetNames = Arr_
End Function

'# Checks if a book path exists
'# This does not only check if it exists, it also checks if it is a book
Public Function CheckBook(BookPath As String) As Boolean
    Dim Book As Workbook
    Set Book = OpenBook(BookPath)
    If Book Is Nothing Then
        CheckBook = False
    Else
        CheckBook = CloseBook(Book)
    End If
End Function

'# Close worksheet safetly
Public Function CloseBook(Book As Workbook, Optional ShouldSave As Boolean = False) As Boolean
    Application.DisplayAlerts = False
    CloseBook = True
    On Error GoTo ErrHandler:
        DoEvents
        DoEvents
        Book.Close SaveChanges:=ShouldSave
        DoEvents
        DoEvents
        Application.DisplayAlerts = True
        Exit Function
ErrHandler:
    Application.DisplayAlerts = True
    CloseBook = False
    Err.Clear
End Function


'# Opens a workbook or bust
Public Function OpenBook(BookPath As String) As Workbook
    Application.DisplayAlerts = False
    On Error GoTo ErrHandler:
        DoEvents
        DoEvents
        Set OpenBook = Workbooks.Open(BookPath)
        DoEvents
        DoEvents
        Application.DisplayAlerts = True
        Exit Function
ErrHandler:
    Application.DisplayAlerts = True
    Set OpenBook = Nothing
    Err.Clear
End Function
