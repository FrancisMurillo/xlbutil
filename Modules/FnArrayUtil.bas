Attribute VB_Name = "FnArrayUtil"
' Functional Array Utility
' ------------------------
'
' These functions support a functional programming style although at the cost of performance
' Primarily these are the familiar map, filter and reduce
' These are added with a suffix of _ to avoid name clashes with any original function and to say they are functionally notated
'
' # Module Contract
'
' MethodNames should be fully qualified as there might be conflict if there is another with the same name.
' Likewise, said methods should follow the argument and return restriction although they are variant
' The type notation is [Arg1, Arg2, ...] -> [Ret], this could be [Var] -> [Int]
'
' # Module Restriction
'
' # Module Dependency
'
' Only FnLambda is required to get the result.
' As an added bonus to use that module as a storage

'# This applies a new array with each element applied to a function
'P MethodName: A function of [Var]->[Var]
'R Retains Base
Public Function Map_(MethodName As String, Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Map_ = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Elem_ As Variant
    Arr_ = ArrayUtil.CloneSize(Arr)

    For Index = LBound(Arr_) To UBound(Arr_)
        Elem_ = Arr(Index)
        Application.Run MethodName, Elem_
        Arr_(Index) = FnLambda.Result
    Next
    
    Map_ = Arr_
End Function

'# This returns a new subarray from an array that satisfies a condition
'P MethodName: A predicate function of [Var]->[Bool], this dictates who gets drafted
'R Zero Base
Public Function Filter_(MethodName As String, Arr As Variant)
    If ArrayUtil.IsEmptyArray(Arr) Then
        Filter_ = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Elem_ As Variant
    Arr_ = ArrayUtil.CreateWithSize(ArrayUtil.Size(Arr))
    For Each Elem_ In Arr
        Application.Run MethodName, Elem_
        If Result Then
            Arr_(Index) = Elem_
            Index = Index + 1
        End If
    Next
    
    If Index = 0 Then
        Arr_ = ArrayUtil.CreateEmptyArray()
    Else
        ReDim Preserve Arr_(0 To Index - 1)
    End If
    
    Filter_ = Arr_
End Function
