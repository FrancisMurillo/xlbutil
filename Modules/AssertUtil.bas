Attribute VB_Name = "AssertUtil"



'# Checks if arrays are first equal in size then by content
'# This follows the ArrayUtil contract for the Empty constant
'# So Empty = Array() in an array point of view
Public Function ArraysEqual(LeftArr As Variant, RightArr As Variant)
    If ArrayUtil.IsEmptyArray(LeftArr) Or ArrayUtil.IsEmptyArray(RightArr) Then
        ArraysEqual = (ArrayUtil.IsEmptyArray(LeftArr) And ArrayUtil.IsEmptyArray(RightArr))
        Exit Function
    End If
    
    ' Check size
    Dim LeftSize As Long, RightSize As Long
    LeftSize = UBound(LeftArr) - LBound(LeftArr)
    RightSize = UBound(RightArr) - LBound(RightArr)
    ArraysEqual = (LeftSize = RightSize)
    If Not ArraysEqual Then Exit Function ' No need to check
    
    ' Check content
    Dim LeftIndex As Long, RightIndex As Long, Index As Long
    LeftIndex = LBound(LeftArr)
    RightIndex = LBound(RightArr)
    While LeftIndex <= UBound(LeftArr)
        ArraysEqual = (LeftArr(LeftIndex) = RightArr(RightIndex))
        If Not ArraysEqual Then Exit Function
        
        LeftIndex = LeftIndex + 1
        RightIndex = RightIndex + 1
    Wend
End Function
