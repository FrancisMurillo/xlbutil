Attribute VB_Name = "TestSheetUtil"
Private gBook As Workbook

Private gSheet1 As Worksheet
Private gSheet2 As Worksheet

Public Sub Setup()
    Set gBook = ActiveWorkbook
    
    Set gSheet1 = gBook.Worksheets.Add
    Set gSheet2 = gBook.Worksheets.Add
End Sub

Public Sub TestDoesSheetExists()
    VaseAssert.AssertTrue _
        SheetUtil.DoesSheetExists(gBook, gSheet1.Name)
    VaseAssert.AssertTrue _
        SheetUtil.DoesSheetExists(gBook, gSheet2.Name)
End Sub

Public Sub TestIsSheetNameUnique()
    Const UNIQUE_NAME As String = "UniqueName" ' This must be guaranteed
    
    VaseAssert.AssertTrue _
        SheetUtil.IsSheetNameUnique(gBook, UNIQUE_NAME), "Unique"

    VaseAssert.AssertFalse _
        SheetUtil.IsSheetNameUnique(gBook, gSheet1.Name), "Sheet1"
    VaseAssert.AssertFalse _
        SheetUtil.IsSheetNameUnique(gBook, gSheet2.Name), "Sheet2"
End Sub

Public Sub TestDeleteSheetSilently()
    Dim ToDeleteSheet As Worksheet, ToDeleteSheetName As String
    Set ToDeleteSheet = gBook.Worksheets.Add
    ToDeleteSheetName = ToDeleteSheet.Name
        
    VaseAssert.AssertTrue _
        SheetUtil.DoesSheetExists(gBook, ToDeleteSheetName)
    SheetUtil.DeleteSheetSilently ToDeleteSheet
    VaseAssert.AssertFalse _
        SheetUtil.DoesSheetExists(gBook, ToDeleteSheetName)
    VaseAssert.AssertErrorNotRaised
End Sub

Public Sub TestAsShortenedSheetName()
    VaseAssert.AssertEqual _
        SheetUtil.AsShortenedSheetName("ShortName"), "ShortName"
    VaseAssert.AssertEqual _
        SheetUtil.AsShortenedSheetName( _
        "AAAAAAAAAABBBBBBBBBCCCCCCCCX"), _
        "AAAAAAAAAABBBBBBBBBCCCCCCCCX"
    VaseAssert.AssertEqual _
        SheetUtil.AsShortenedSheetName( _
        "AAAAAAAAAABBBBBBBBBCCCCCCCCXABCY"), _
        "AAAAAAAAAABBBBBBBBBCCCCCCCCX..."
End Sub

Public Sub TestStripBadCharacters()
    VaseAssert.AssertEqual _
        SheetUtil.StripBadCharacters("CREATE: More sheets"), _
        "CREATE More sheets"
    VaseAssert.AssertEqual _
        SheetUtil.StripBadCharacters("Sheet:\/?*[]"), _
        "Sheet"
End Sub

' This also test GetLastSheet, MoveSheetToEnd
Public Sub TestRemoveAllSheetsExceptFirstFew()
    Const SHEET_COUNT As Integer = 10
    
    Dim LastSheet As Worksheet, TempSheet As Worksheet, Index As Integer
    Dim Count As Integer
    Set LastSheet = SheetUtil.GetLastSheet(gBook)
    Count = gBook.Worksheets.Count
        
    For Index = 1 To SHEET_COUNT
        Set TempSheet = gBook.Worksheets.Add
        SheetUtil.MoveSheetToEnd gBook, TempSheet
    Next
    
    SheetUtil.RemoveAllSheetExceptFirstFew gBook, LastSheet.Index
    
    VaseAssert.AssertEqual Count, gBook.Worksheets.Count
    
End Sub

Public Sub Teardown()
    SheetUtil.DeleteSheetSilently gSheet1
    SheetUtil.DeleteSheetSilently gSheet2
End Sub
