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
