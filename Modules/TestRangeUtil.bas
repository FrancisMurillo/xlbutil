Attribute VB_Name = "TestRangeUtil"
Private gBook As Workbook
Private gSheet As Worksheet
Private gRange As Range

Public Sub Setup()
    Set gBook = ActiveWorkbook
    Set gSheet = gBook.Worksheets.Add
    
    With gSheet
        .Cells(2, 2).Value = 1
        .Cells(2, 3).Value = 2
        .Cells(2, 4).Value = 3
        
        .Cells(3, 2).Value = 2
        .Cells(3, 3).Value = 3
        .Cells(3, 4).Value = 4
        
        Set gRange = gSheet.Range(.Cells(2, 2), .Cells(3, 4))
    End With
End Sub

Public Sub Teardown()
    SheetUtil.DeleteSheetSilently gSheet
End Sub

Public Sub TestUpperLeftCell()
    Dim ExpectedCell As Range, ActualCell As Range
    Set ExpectedCell = gSheet.Cells(2, 2)
    Set ActualCell = RangeUtil.GetUpperLeftCell(gRange)
    
    VaseAssert.AssertTrue CellEq_(ActualCell, ExpectedCell)
End Sub

Public Sub TestUpperRightCell()
    Dim ExpectedCell As Range, ActualCell As Range
    Set ExpectedCell = gSheet.Cells(2, 4)
    Set ActualCell = RangeUtil.GetUpperRightCell(gRange)
    
    VaseAssert.AssertTrue CellEq_(ActualCell, ExpectedCell)
End Sub

Public Sub TestLowerLeftCell()
    Dim ExpectedCell As Range, ActualCell As Range
    Set ExpectedCell = gSheet.Cells(3, 2)
    Set ActualCell = RangeUtil.GetLowerLeftCell(gRange)
    
    VaseAssert.AssertTrue CellEq_(ActualCell, ExpectedCell)
End Sub

Public Sub TestLowerRightCell()
    Dim ExpectedCell As Range, ActualCell As Range
    Set ExpectedCell = gSheet.Cells(3, 4)
    Set ActualCell = RangeUtil.GetLowerRightCell(gRange)
    
    VaseAssert.AssertTrue CellEq_(ActualCell, ExpectedCell)
End Sub

Public Sub TestAsArray()
    Dim RowArr As Variant, ColArr As Variant
    Dim ExRowArr As Variant, ExColArr As Variant
    
    RowArr = RangeUtil.AsRowArray(gRange, 1)
    ColArr = RangeUtil.AsColumnArray(gRange, 1)
    ExRowArr = Array(1, 2, 3)
    ExColArr = Array(1, 2)
    VaseAssert.AssertArraysEqual _
        RowArr, ExRowArr
    VaseAssert.AssertArraysEqual _
        ColArr, ExColArr
    
    
    RowArr = RangeUtil.AsRowArray(gRange, 2)
    ColArr = RangeUtil.AsColumnArray(gRange, 2)
    ExRowArr = Array(2, 3, 4)
    ExColArr = Array(2, 3)
    VaseAssert.AssertArraysEqual _
        RowArr, ExRowArr
    VaseAssert.AssertArraysEqual _
        ColArr, ExColArr
    
End Sub


Public Function TestAsColumnArrays()
    Setup
    Rewind_
    
    Dim Values As Variant
    Values = RangeUtil.AsColumnArrays(gRange, Array(1, 3))
    
    VaseAssert.AssertTrue True
    Ping_
End Function

Private Function CellEq_(LeftCell As Range, RightCell As Range) As Boolean
    CellEq_ = (LeftCell.Row = RightCell.Row) And _
                (LeftCell.Column = RightCell.Column) And _
                (LeftCell.Value = RightCell.Value)
End Function


