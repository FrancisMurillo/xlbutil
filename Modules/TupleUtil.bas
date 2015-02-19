Attribute VB_Name = "TupleUtil"
'===========================
'--- Module Contract     ---
'===========================
'# MODULE DEFINITION
'# Tuples are the superset of Arrays
'# They are arrays, they can be arrays of arrays.
'# This super structure allow nested tuples and better handling of them.
'#
'# MODULE LIMITATION
'# This definition is not performance friendly, so do not expect to use this on performance intensive application
'#
'# MODULE CONTRACT
'# Tuples follow these rules
'# 1. Data Types - Strings, Numbers, Dates, Tuples but NOT Objects(wrap it in a tuple)
'# 2. Tuples are zero-indexed,

'# This transposes the rectangular tuples like matrix transposition
'! Assumes TArr is a rectangular array or elements are arrays and are of equal sizes else this raises a runtime error
Public Function Transpose(TArr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(TArr) Then
        TransposeTuples = Array()
        Exit Function
    End If
    
    Dim TArr_ As Variant
    Dim LeftArr As Variant, RightArr As Variant, Tuple As Variant, FirstElement As Variant
    TArr_ = ArrayUtil.CloneSize(TArr)
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

'===========================
'--- Operator            ---
'===========================

'# Compares tuples
'# This equates Empty and the empty array as the same since they are IsEmptyArray
Public Function EqualTuple_(LeftTuple As Variant, RightTuple As Variant) As Boolean
On Error Resume Next
    If ArrayUtil.IsEmptyArray(LeftTuple) Then
        EqualTuple_ = ArrayUtil.IsEmptyArray(RightTuple)
    ElseIf IsArray(LeftTuple) Then
    
    Else
        If IsArray(RightTuple) Then
        
        End If
    End If
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
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


