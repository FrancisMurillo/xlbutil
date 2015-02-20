Attribute VB_Name = "FnArrayUtil"
' Functional Array Utility
' ------------------------
'
' These functions support a functional programming style although at the cost of performance
' Primarily these are the familiar map, filter and reduce
'
' # Module Contract
'
' MethodNames should be fully qualified as there might be conflict if there is another with the same name.
'
' # Module Restriction
'
' # Module Dependency
'
' Only FnLambda is required to get the result.
' As an added bonus to use that module as a storage

'P MethodName: A function
Public Function Map(MethodName As String, Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Map = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Elem_ As Variant
    Arr_ = ArrayUtil.CloneSize(Arr)

    For Index = LBound(Arr_) To UBound(Arr_)
        Elem_ = Arr(Index)
        Application.Run MethodName, Elem_
        Arr_(Index) = FnLambda.Result
    Next
    
    Map = Arr_
End Function
