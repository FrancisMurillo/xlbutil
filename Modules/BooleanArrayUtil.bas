Attribute VB_Name = "BooleanArrayUtil"
' Boolean Array Util
' ------------------
'
' Another set of utilities of boolean arrays

'# Returns true if any of the elements are true
'# If the array is empty, do it like Python, False
Public Function Any_(BArr As Variant) As Boolean
    Any_ = False
    If ArrayUtil.IsEmptyArray(BArr) Then _
        Exit Function
        
    Dim Elem_ As Variant
    For Each Elem_ In BArr
        Any_ = Any_ Or Elem_
        If Any_ Then Exit Function
    Next
End Function

'# Checks if all elements are true
'# If the array is empty, do it like Python, False
Public Function All_(BArr As Variant) As Boolean
    All_ = False
    If ArrayUtil.IsEmptyArray(BArr) Then _
        Exit Function
    All_ = True
        
    Dim Elem_ As Variant
    For Each Elem_ In BArr
        All_ = All_ And Elem_
        If Not All_ Then Exit Function
    Next
End Function


