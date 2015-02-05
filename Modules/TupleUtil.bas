Attribute VB_Name = "TupleUtil"
'===========================
'--- Module Contract     ---
'===========================
' Tuples are like Arrays but still under contention on how they function
' Primarily, tuples are arrays that can have arrays as elements which differs from ArrayUtil.

'# This just splits a tuple array into a larger array of one tuple
'# A transposition of sorts
'# This follows the ArrayUtil contract
Public Function TransposeTuples(Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        TransposeTuples = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant
    Dim LeftArr As Variant, RightArr As Variant, Tuple As Variant, FirstElement As Variant
    Arr_ = ArrayUtil.CloneSize(Arr)
    FirstElement = Arr(0)
    Tuples = ArrayUtil.CloneSize(FirstElement)
    
    ' Setup the array
    Dim I As Long, J As Long
    For I = 0 To UBound(FirstElement)
        Tuples(I) = ArrayUtil.CloneSize(Arr)
    Next
    ' Fill it up
    For I = 0 To UBound(Arr)
        For J = 0 To UBound(Arr(I))
            Tuples(J)(I) = Arr(I)(J)
        Next
    Next
    
    TransposeTuples = Tuples
End Function

'# Sorts an array of tuples based on its index
'# This used the primitive bubble sort
'! If there are errors in value sorting, it will not be cleared
Public Function SortTuples(TupArr As Variant, SortIndex As Long, Optional SortingOption As TupleSorting = TupleSorting.ASCENDING) As Variant
    If ArrayUtil.IsEmptyArray(TupArr) Then
        SortTuples = Array()
        Exit Function
    End If
    
    Dim TArr_ As Variant, I As Long, J As Long, Tmp As Variant
    Dim Left As Variant, Right As Variant
    TArr_ = TupArr
    For I = 0 To UBound(TArr_)
        For J = 1 To UBound(TArr_)
            Left = TArr_(J - 1)(SortIndex)
            Right = TArr_(J)(SortIndex)
            If SortingOption = ASCENDING And GreaterThan_(Left, Right) Then
                Tmp = TArr_(J - 1)
                TArr_(J - 1) = TArr_(J)
                TArr_(J) = Tmp
            ElseIf SortingOption = DESCENDING And LessThan_(Left, Right) Then
                Tmp = TArr_(J - 1)
                TArr_(J - 1) = TArr_(J)
                TArr_(J) = Tmp
            End If
        Next
    Next
    SortTuples = TArr_
End Function

'# This removes all empty tuples in an array of tuples
'# This is different from RemoveAllDuplicates of ArrayUtil as this handles Empty Arrays specifically
Public Function RemoveAllEmptyTuples(TArr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(TArr) Then
        RemoveAllEmptyTuples = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Index As Long, Item As Variant, ArrIndex As Long
    Arr_ = CloneSize(TArr)
    ArrIndex = 0
    For Index = 0 To UBound(TArr)
        Item = TArr(Index)
        If Not ArrayUtil.IsEmptyArray(Item) Then
            Arr_(ArrIndex) = Item
            ArrIndex = ArrIndex + 1
        End If
    Next
    
    If ArrIndex = 0 Then ' None was left
        Arr_ = Array()
    Else
        ReDim Preserve Arr_(0 To ArrIndex - 1)
    End If
    
    RemoveAllEmptyTuples = Arr_
End Function

'# This filters tuples based on a certain value and index
Public Function FilterTuples(Elem As Variant, ColIndex As Long, TArr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(TArr) Then
        FilterTuples = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Index As Long, Item As Variant, ArrIndex As Long
    Arr_ = CloneSize(TArr)
    ArrIndex = 0
    For Index = 0 To UBound(TArr)
        Item = TArr(Index)
        If Equal_(Item(ColIndex), Elem) Then
            Arr_(ArrIndex) = Item
            ArrIndex = ArrIndex + 1
        End If
    Next
    
    If ArrIndex = 0 Then ' None was left
        Arr_ = Array()
    Else
        ReDim Preserve Arr_(0 To ArrIndex - 1)
    End If
    
    FilterTuples = Arr_
End Function

