Attribute VB_Name = "RangeArrayUtil"
' Range Array Utilities
' ---------------------
'
' Utilities to using Range as Arrays, this reasoning is for performance and all

'# Returns an array representing the range values
Public Function ToArray(Rng As Range) As Variant
    Dim RngArr As Variant
    RngArr = Array()
    ReDim RngArr(1 To Rng.Rows.CountLarge, 1 To Rng.Columns.CountLarge)
    
    Rng.Value
End Function
