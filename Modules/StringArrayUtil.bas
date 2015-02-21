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

'# This checks if a string is within a string array with the like operator
'C No Zero Base Restriction
Public Function IsInLike(Pattern As String, SArr As Variant, _
                    Optional IgnoreCase As Boolean = False) As Boolean
    IsInLike = False
    If ArrayUtil.IsEmptyArray(SArr) Then _
        Exit Function

    Dim Match As Variant, SMatch As String, Pattern_
    Pattern_ = Pattern
    If IgnoreCase Then _
        Pattern_ = UCase(Pattern)
        
    For Each Match In SArr
        SMatch = Match
        If IgnoreCase Then _
            SMatch = UCase(SMatch)
            
        If SMatch Like Pattern_ Then
            IsInLike = True
            Exit Function
        End If
    Next
End Function

'# This complements IsInLike as this gives the index found.
'# However, this assumes the array is zero base
'# since an incremental find would not be possible without a fixed base
'P StartIndex: This dictates where to start the search, you can use this to search for all the elements with this criteria
'C Base Indepedent
'R If it is not found, it gives one less than the start index.
'R If it is found, it gives the first index with that index
Public Function FindLike(Pattern As String, SArr As Variant, _
                    Optional StartIndex As Long = 0, _
                    Optional IgnoreCase As Boolean = False) As Long
    FindLike = -1
    If ArrayUtil.IsEmptyArray(SArr) Then _
        Exit Function
    
    Dim Index As Long, Pattern_ As String, SMatch As String
    Pattern_ = Pattern
    If IgnoreCase Then _
        Pattern_ = UCase(Pattern)
    
    For Index = StartIndex To UBound(SArr)
        SMatch = SArr(Index)
        If IgnoreCase Then _
            SMatch = UCase(SMatch)
            
        If SMatch Like Pattern_ Then
            FindLike = Index
            Exit Function
        End If
    Next
End Function
