Attribute VB_Name = "ArrayUtil"
'===========================
'--- Module Contract     ---
'===========================
' All arrays are...
' 1. Zero-based - Arrays should start at zero. Use ShiftBase() to make an array start at zero
' 2. Variants - Arrays should be easy to type
' 3. Single Dimension - Multi-dimensional arrays should be just array of arrays
' That is how arrays should be
' -----
' What they can be are.
' 1. Empty, the constant. The case returns an empty array usually unless mentioned
' 2. An array with no elements
' -----
' Other contracts are
' 1. Integers/Longs are non-negative unless specified
' 2. Exceptions are raised at this level, they are not handled because they are handled.
' 3. No objects! Just plain Variants

'===========================
'--- Constants           ---
'===========================

'# Some error object constants
Private Const ERR_SOURCE As String = "ArrayUtil"
Private Const ERR_OFFSET As Long = 1000

'===========================
'--- Core Functions      ---
'===========================

'# Re-shifts an array base to a specified start index, the ideal start index is 0 but the option is there
Public Function ShiftBase(Arr As Variant, Optional Index As Long = 0) As Variant
    Dim Arr_ As Variant
    Arr_ = AsArray(Arr)
    If LBound(Arr_) = 0 Then
        ShiftBase = Arr_
        Exit Function ' Array has been normalized
    End If
    ReDim Preserve Arr_(0 To UBound(Arr_) - LBound(Arr_)) 'Shift base
    ShiftBase = Arr_
End Function

'# This is basically reframes ShiftBase() as the array normalizer for this library
Public Function AsNormalArray(Arr As Variant) As Variant
    AsNormalArray = ShiftBase(Arr, Index:=0)
End Function

'# Checks if an array is Empty or has no elements
Public Function IsEmptyArray(Arr As Variant) As Boolean
    IsEmptyArray = False ' Default to False
    If IsEmpty(Arr) Then ' Handles Empty constant
        IsEmptyArray = True
    ElseIf IsArray(Arr) Then
        IsEmptyArray = (UBound(Arr) < LBound(Arr)) ' The empty array condition
    Else ' Not supposed to happen, this might be a type error
        Err.Raise vbObjectError + ERR_OFFSET, Source:=ERR_SOURCE, Description:="IsEmptyArray"
    End If
End Function

'# Turns an empty array, as defined by IsEmptyArr, to an empty array
'# This is basically a utility to avoid the if condition and just return a default empty array
'# Otherwise it returns the array passed into it
Public Function AsArray(Arr As Variant) As Variant
    AsArray = IIf(IsEmptyArray(Arr), Array(), Arr)
End Function

'# This returns an array as the same size as the array
'! Input array can be of any base
Public Function CloneSize(Arr As Variant) As Variant
    If IsEmptyArray(Arr) Then
        CloneSize = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant
    Arr_ = Array()
    ReDim Arr_(LBound(Arr) To UBound(Arr))
    CloneSize = Arr_
End Function

'# This returns the size of an array
'! Input array can be of any base
Public Function Size(Arr As Variant) As Long
    Size = 0
    If IsEmpty(Arr) Then Exit Function
    
    Size = UBound(Arr) - LBound(Arr) + 1
End Function


'===========================
'--- Secondary Functions ---
'===========================
' Majority of the functions here assume the zero-base rule

'# Removes all elements in an array that matches the one specified
Public Function RemoveAllElements(Elem As Variant, Arr As Variant) As Variant
    If IsEmptyArray(Arr) Then
        RemoveAllElements = Array()
        Exit Function
    End If
    
    
    Dim Arr_ As Variant, Index As Long, Item As Variant, ArrIndex As Long
    Arr_ = CloneSize(Arr)
    ArrIndex = 0
    For Index = 0 To UBound(Arr)
        Item = Arr(Index)
        If Not TryEqual(Elem, Item) Then
            Arr_(ArrIndex) = Item
            ArrIndex = ArrIndex + 1
        End If
    Next
    
    If ArrIndex = 0 Then ' None was left
        Arr_ = Array()
    Else
        ReDim Preserve Arr_(0 To ArrIndex - 1)
    End If
    
    RemoveAllElements = Arr_
End Function

'# Same as RemoveAllElements except it is tuned against Empty elements
Public Function RemoveAllEmptyElements(Arr As Variant) As Variant
    RemoveAllEmptyElements = RemoveAllElements(Empty, Arr)
End Function

'# Checks if two elements are equal without the type error noise.
'# Defaults to False on error and clears the error
Private Function TryEqual(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim HasErrorAlready As Boolean
    HasErrorAlready = (Err.Number <> 0)
On Error Resume Next
    TryEqual = False
    TryEqual = (LeftVal = RightVal)
    If Not HasErrorAlready Then Err.Clear ' Clear the error if there was none at the start
End Function
