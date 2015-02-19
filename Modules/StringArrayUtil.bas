Attribute VB_Name = "StringArrayUtil"
' Module Contract
' ---------------
' This follows the same contract as ArrayUtil with the every element being a string
' Quite obvious given the name of the module

'# This trims every string in the array, a cleansing function basically
'C No Zero Based Restriction
'@ Return: The array with every element trimmed
Public Function TrimAll(SArr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(SArr) Then
        TrimAll = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Index As Long
    Arr_ = ArrayUtil.CloneSize(SArr)
    For Index = LBound(SArr) To UBound(SArr)
        Arr_(Index) = Trim(SArr(Index))
    Next
    
    TrimAll = Arr_
End Function

