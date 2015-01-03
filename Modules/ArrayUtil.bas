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
    If IsEmptyArray(Arr) Then
        ShiftBase = Array()
        Exit Function
    End If

    Dim Arr_ As Variant, Offset As Long, Lower As Long, Upper As Long
    Arr_ = AsArray(Arr)
    
    Lower = LBound(Arr_)
    Upper = UBound(Arr_)
    
    Offset = Index - Lower
    
    ReDim Preserve Arr_(Lower + Offset To Upper + Offset)  'Shift base
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
Public Function IsNotEmptyArray(Arr As Variant) As Boolean
    IsNotEmptyArray = Not IsEmptyArray(Arr)
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
        If Not Equal_(Elem, Item) Then
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
Private Function Equal_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim HasErrorAlready As Boolean
    HasErrorAlready = (Err.Number <> 0)
On Error Resume Next
    Equal_ = False
    Equal_ = (LeftVal = RightVal)
    If Not HasErrorAlready Then Err.Clear ' Clear the error if there was none at the start
End Function


'# Removes every duplicate element in an array
'# This assumes the array is homogenous or risk having errors
Public Function RemoveDuplicates(Arr As Variant) As Variant
    If IsEmptyArray(Arr) Then ' Empty array check
        RemoveDuplicates = Array()
        Exit Function
    End If
    
    ' Return value setup
    Dim Arr_ As Variant
    Arr_ = CloneSize(Arr)
    
    ' Loop through duplicates
    Dim Index As Long, Count As Long, Item As Variant
    Count = 0
    For Index = 0 To UBound(Arr)
        Item = Arr(Index)
        If Not IsIn(Item, Arr_) Then ' Check if item is not in the pseudo set then add it
            Arr_(Count) = Item
            Count = Count + 1
        End If
    Next
    
    If Count > 0 Then ' Resize the array
        ReDim Preserve Arr_(0 To Count - 1)
    Else ' Empty array check again
        SetArr = Array()
    End If
    
    RemoveDuplicates = Arr_
End Function

'# Checks if an element is in an array
Public Function IsIn(Elem As Variant, Arr As Variant) As Variant
    IsIn = False
    Dim Item As Variant
    For Each Item In AsArray(Arr)
        IsIn = Equal_(Item, Elem)
        If IsIn Then Exit Function
    Next
End Function

'# Checks if an array has duplicate elements
Public Function HasDuplicates(Arr As Variant) As Variant
    If IsEmptyArray(Arr) Then ' Empty array check
        HasDuplicates = False
        Exit Function
    End If

    Dim I As Long, J As Long
    For I = 0 To UBound(Arr)
        For J = I + 1 To UBound(Arr)
            If Equal_(Arr(I), Arr(J)) Then
                HasDuplicates = True
                Exit Function
            End If
        Next
    Next
End Function

'# Joins two arrays into one array
'# This accepts two zero-based arrays
Public Function JoinArrays(LeftArr As Variant, RightArr As Variant) As Variant
    Dim IsLeftEmpty As Boolean, IsRightEmpty As Boolean
    IsLeftEmpty = IsEmptyArray(LeftArr)
    IsRightEmpty = IsEmptyArray(RightArr)
    
    If IsLeftEmpty And IsRightEmpty Then
        JoinArrays = Array()
        Exit Function
    ElseIf IsLeftEmpty Or IsRightEmpty Then
        JoinArrays = IIf(IsLeftEmpty, RightArr, LeftArr)
        Exit Function
    Else
        Dim Arr_ As Variant, Index As Long
        Arr_ = Array()
        ReDim Arr_(0 To (UBound(LeftArr) + UBound(RightArr) + 1))
        For Index = 0 To UBound(Arr_)
            If Index <= UBound(LeftArr) Then
                Arr_(Index) = LeftArr(Index)
            Else
                Arr_(Index) = RightArr(Index - UBound(LeftArr) - 1)
            End If
        Next
        JoinArrays = Arr_
    End If
End Function


'# Creates an range array similar to Python
'# Used in creating the mapping values
Public Function Range(Optional Start_ As Long = 0, Optional Stop_ As Long = 0, Optional Step_ As Long = 1)
    If Start_ = Stop_ Then ' A one-element range
        Range = Array(Start_)
        Exit Function
    End If
    
    If Step_ = 0 Then ' Highly unusual but this is an error
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Range"
    ElseIf Stop_ > Start_ And Not Step_ > 0 Then ' The direction must match the step
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Range"
    ElseIf Stop_ < Start_ And Not Step_ < 0 Then ' The direction must match the step
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Range"
    End If
    
    Dim Index As Long, Counter As Long
    Dim Size As Long, Rng_ As Variant
    Counter = Start_
    If Step_ > 0 Then
        Size = CLng(Abs(((Stop_ - 1) - Start_) / Step_))
    Else
        Size = CLng(Abs(((Start_ - 1) - Stop_) / Step_))
    End If

    
    Rng_ = Array()
    ReDim Rng_(0 To Size)
    For Index = 0 To UBound(Rng_)
        Rng_(Index) = Counter
        Counter = Counter + Step_
    Next
    
    Range = Rng_
End Function


'# This is a helper function that creates an array with the designated size
'# Size is assumed to be greater than 1 or get an empty array
Public Function CreateWithSize(Size As Long) As Variant
    If Size < 1 Then
        CreateWithSize = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant
    Arr_ = Array()
    ReDim Arr_(0 To Size - 1)
    CreateWithSize = Arr_
End Function


'# This function returns a subarray of an array by giving its indices
'# This is like a SELECT for Array
'! If there is an index that is out of bounds, this will throw that error
Public Function Projection(Indices As Variant, Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Indices) Or ArrayUtil.IsEmptyArray(Arr) Then
        Projection = Array()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Index As Long
    Arr_ = CloneSize(Indices)
    
    For Index = 0 To UBound(Indices)
        Arr_(Index) = Arr(Indices(Index))
    Next
    
    Projection = Arr_
End Function

'# Creates a set difference with a rudimentary arrays
Public Function SetDifference(Subset_ As Variant, Set_ As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Subset_) Or ArrayUtil.IsEmptyArray(Set_) Then
        SetDifference = Array()
        Exit Function
    End If
    
    Dim Index As Long, Count As Long, Elem As Variant, ArrSet_ As Variant
    ArrSet_ = CloneSize(Set_)
    For Index = 0 To UBound(Set_)
        Elem = Set_(Index)
        If Not IsIn(Elem, Subset_) Then
            ArrSet_(Count) = Elem
            Count = Count + 1
        End If
    Next
    
    If Count = 0 Then
        ArrSet_ = Array()
    Else
        ReDim Preserve ArrSet_(0 To Count - 1)
    End If
    
    SetDifference = ArrSet_
End Function

'# Creates an empty array
Public Function CreateEmptyArray() As Variant
    CreateEmptyArray = Array()
End Function
