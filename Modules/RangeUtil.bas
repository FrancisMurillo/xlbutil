Attribute VB_Name = "RangeUtil"
'===========================
'--- Module Contract     ---
'===========================
' This is for Range objects

'# The two pairs of function gets the last row and column respectively
Public Function GetLastRow(Rng As Range) As Range
    Set GetLastRow = Rng.Rows(Rng.Rows.CountLarge)
End Function
Public Function GetLastCol(Rng As Range) As Range
    Set GetLastCol = Rng.Columns(Rng.Columns.CountLarge)
End Function

'# Another set to reflect the first row and column
Public Function GetFirstRow(Rng As Range) As Range
    Set GetFirstRow = Rng.Rows(1)
End Function
Public Function GetFirstCol(Rng As Range) As Range
    Set GetFirstCol = Rng.Columns(1)
End Function

'# These four takes the upper and lower most sides
Public Function GetUpperLeftCell(Rng As Range) As Range
    Set GetUpperLeftCell = RangeUtil.GetFirstCol( _
                            RangeUtil.GetFirstRow(Rng))
End Function
Public Function GetUpperRightCell(Rng As Range) As Range
    Set GetUpperRightCell = RangeUtil.GetLastCol( _
                            RangeUtil.GetFirstRow(Rng))
End Function
Public Function GetLowerLeftCell(Rng As Range) As Range
    Set GetLowerLeftCell = RangeUtil.GetFirstCol( _
                            RangeUtil.GetLastRow(Rng))
End Function
Public Function GetLowerRightCell(Rng As Range) As Range
    Set GetLowerRightCell = RangeUtil.GetLastCol( _
                            RangeUtil.GetLastRow(Rng))
End Function


'# These function gets the values of an row or column or both as an array
Public Function AsRowArray(Rng As Range, RowIndex As Long) As Variant
    Dim Row As Range, Arr As Variant, Index As Long
    Set Row = Rng.Rows(RowIndex)
    Arr = Array()
    ReDim Arr(0 To Row.Columns.Count - 1)
    For Index = 0 To UBound(Arr)
        Arr(Index) = Row.Columns(Index + 1).Value
    Next
    
    AsRowArray = Arr
End Function
Public Function AsColumnArray(Rng As Range, ColIndex As Long) As Variant
    Dim Col As Range, Arr As Variant, Index As Long
    Set Col = Rng.Columns(ColIndex)
    Arr = Array()
    ReDim Arr(0 To Col.Rows.Count - 1)
    For Index = 0 To UBound(Arr)
        Arr(Index) = Col.Rows(Index + 1).Value
    Next
    
    AsColumnArray = Arr
End Function

'# Extended function of AsColumnArray where an array of indices are given
'# This is tweaked for some performance aspect
Public Function AsColumnArrays(Rng As Range, ColIndices As Variant) As Variant
    If ArrayUtil.IsEmptyArray(ColIndices) Then
        AsColumnArrays = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Cols_ As Variant, Indices_ As Variant, Index As Long, ColIndex As Long
    Arr_ = ArrayUtil.ShiftBase(ArrayUtil.CreateWithSize(RangeUtil.GetRowCount(Rng)), 1)
    Cols_ = ArrayUtil.ShiftBase(ArrayUtil.CloneSize(ColIndices), 1)
    Indices_ = ArrayUtil.ShiftBase(ColIndices, 1)
    For Index = 1 To UBound(Cols_)
        ColIndex = Indices_(Index)
        Cols_(Index) = AppUtil.WrapRangeAsArray(Rng.Columns(ColIndex).Value)
    Next
    
    Dim TempArr_ As Variant, RowIndex As Long, TempIndex As Long, ColSize As Long
    ColSize = ArrayUtil.Size(ColIndices)
    TempArr_ = ArrayUtil.CloneSize(Cols_)
    For Index = 1 To UBound(Arr_)
        For TempIndex = 1 To ColSize ' Collection Array, starts at 1
            TempArr_(TempIndex) = Cols_(TempIndex)(Index, 1)
        Next
        Arr_(Index) = TempArr_
    Next
    
    AsColumnArrays = Arr_
End Function


'# A pair of function to get row and column count from a range
Public Function GetRowCount(Rng As Range) As Long
    GetRowCount = Rng.Rows.CountLarge
End Function
Public Function GetColumnCount(Rng As Range) As Long
    GetColumnCount = Rng.Columns.CountLarge
End Function


