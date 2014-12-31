Attribute VB_Name = "SheetUtil"
'===========================
'--- Module Contract     ---
'===========================
' This encompases both Workbook and Worksheet utilities
' Sheet and Worksheet are synonyms, and so are Book and Workbook
' There is one caveat when creating sheets in Excel
' * Sheet Names can only have 31 characters at most
'       Functions handle it by first triming it, then giving it another name until satisfied



'===========================
'--- Constants           ---
'===========================
Public Const SHEET_NAME_LENGTH_LIMIT As Integer = 31
Public Const ELLIPSIS As String = "..."

Private Const BAD_SHEET_NAME_CHARACTERS As String = ":;\;/;?;*;[;]"

'===========================
'--- Properties           ---
'===========================

Public Property Get BadSheetNameCharacters() As Variant
    BadSheetNameCharacters = Split(BAD_SHEET_NAME_CHARACTERS, ";")
End Property

'===========================
'--- Functions     ---
'===========================

'# Checks if an worksheet exists
'? Reference: http://www.mrexcel.com/forum/excel-questions/3228-visual-basic-applications-check-if-worksheet-exists.html
Public Function DoesSheetExists(Book As Workbook, SheetName As String) As Boolean
    DoesSheetExists = False
    
    Dim Index As Integer, Sheet As Worksheet
    For Index = 1 To Book.Sheets.Count
        Set Sheet = Book.Worksheets(Index)
        DoesSheetExists = (Sheet.Name = SheetName)
        If DoesSheetExists Then Exit Function
    Next
End Function

'# Produce a sheet name that doesn't exceed the character limit
Public Function AsShortenedSheetName(SheetName As String, Optional Filler As String = ELLIPSIS) As String
    If Len(SheetName) <= SHEET_NAME_LENGTH_LIMIT Then
        AsShortenedSheetName = SheetName
    Else
        AsShortenedSheetName = Left(SheetName, SHEET_NAME_LENGTH_LIMIT - Len(Filler)) & Filler
    End If
End Function

'# This strips the bad characters in a sheet
'# These bad characters are : \ / ? * [ ]
Public Function StripBadCharacters(SheetName As String, Optional Replacement As String = "") As String
    Dim StrippedName As String, Char As Variant
    StrippedName = SheetName
    
    For Each Char In BadSheetNameCharacters
        StrippedName = Replace(StrippedName, CStr(Char), Replacement)
    Next
    
    StripBadCharacters = StrippedName
End Function

'# This checks if a sheet name is unique in a book.
'# This is used for safely renaming a sheet
'! SheetName however should comply with the 31 character rule
'!  This just checks the name, handling it is someone elses work
Public Function IsSheetNameUnique(Book As Workbook, SheetName As String) As Boolean
    IsSheetNameUnique = Not DoesSheetExists(Book, SheetName)
End Function

'# Quietly delete a sheet without the clutter
'# Used in programatic deletes without the user interaction
Public Sub DeleteSheetSilently(Sheet As Worksheet)
On Error Resume Next
    Dim HasNoError As Boolean
    HasNoError = (Err.Number = 0)
    
    Application.DisplayAlerts = False
    Sheet.Delete
    Application.DisplayAlerts = True
    
    If HasNoError Then Err.Clear
End Sub

'# This removes all sheets in a workbook except the first N, where N is an integer
'# This also assumes the sheets are already ordered to what gets deleted and not
'@ Param(Count): This assumes N is greater than 0, else it does nothing
'@              So if this is 1, this removes all sheet but the first
Public Sub RemoveAllSheetExceptFirstFew(Book As Workbook, Count As Integer)
    Application.DisplayAlerts = False ' Remove alerts for deleting a sheet
    Do While ActiveWorkbook.Sheets.Count > Count
        ActiveWorkbook.Sheets(Count + 1).Delete
    Loop
    Application.DisplayAlerts = True
End Sub

'# Gets the last sheet in a workbook
Public Function GetLastSheet(Book As Workbook)
    Set GetLastSheet = Book.Worksheets(Book.Worksheets.Count)
End Function

'# Moves a sheet to the end, nothing fancy
Public Sub MoveSheetToEnd(Book As Workbook, Sheet As Worksheet)
    Sheet.Move After:=GetLastSheet(Book)
End Sub

'# A pair of function to count there records
Public Function GetRowCount(Sheet As Worksheet) As Long
    GetRowCount = Sheet.UsedRange.Rows.CountLarge
End Function
Public Function GetColumnCount(Sheet As Worksheet) As Long
    GetColumnCount = Sheet.UsedRange.Rows.CountLarge
End Function
