Attribute VB_Name = "TestArrayUtil"
' Testing Array Util
' Test for these things at least
' 1. Empty Array
' 2. Empty Constant

Private gEmptyConstant As Variant
Private gEmptyArray As Variant
Private gNormalArray As Variant
Private gNonNormalArray As Variant
Private gNormalizedArray As Variant
Private gRandomValue As Variant

Private Sub Setup()
    gEmptyArray = Array()
    gEmptyConstant = Empty
    gNormalArray = Array(1, "1", "One")
    gNonNormalArray = Array(1, 2, 3)
    gNormalizedArray = gNonNormalArray ' Before mutating base
    ReDim Preserve gNonNormalArray(3 To 5)
    
    gRandomValue = Now
End Sub

Public Sub TestIsEmptyArray()
    VaseAssert.AssertTrue ArrayUtil.IsEmptyArray(gEmptyArray), "Empty Array"
    VaseAssert.AssertTrue ArrayUtil.IsEmptyArray(gEmptyConstant), "EMPTY"
    VaseAssert.AssertFalse ArrayUtil.IsEmptyArray(gNormalArray), "Normal Array"
    VaseAssert.AssertFalse ArrayUtil.IsEmptyArray(gNonNormalArray), "Shifted Array"
    
    VaseAssert.AssertFalse ArrayUtil.IsEmptyArray(Array(1)) ' Test single value
End Sub

Public Sub TestAsArray()
    VaseAssert.AssertArraysEqual ArrayUtil.AsArray(gEmptyArray), Array(), "Empty Array"
    VaseAssert.AssertArraysEqual ArrayUtil.AsArray(gEmptyConstant), Array(), "EMPTY"
    VaseAssert.AssertArraysEqual ArrayUtil.AsArray(gNormalArray), gNormalArray, "Normal Array"
    VaseAssert.AssertArraysEqual ArrayUtil.AsArray(gNonNormalArray), gNonNormalArray, "Shifted Array"
End Sub

Public Sub TestAsNormalArray()
    VaseAssert.AssertArraysEqual ArrayUtil.AsNormalArray(gEmptyArray), Array(), "Empty Array"
    VaseAssert.AssertArraysEqual ArrayUtil.AsNormalArray(gEmptyConstant), Array(), "EMPTY"
    VaseAssert.AssertArraysEqual ArrayUtil.AsNormalArray(gNormalArray), gNormalArray, "Normal Array"
    VaseAssert.AssertArraysEqual ArrayUtil.AsNormalArray(gNonNormalArray), gNonNormalArray, "Shifted Array"
End Sub

Public Sub TestShiftBase()
    Dim TempArr As Variant, OutArr As Variant
    VaseAssert.AssertArraysEqual gNormalizedArray, gNonNormalArray, "Shifted array"
    
    TempArr = ArrayUtil.ShiftBase(gNormalizedArray, 3)
    VaseAssert.AssertEqual LBound(TempArr), 3, "Normalized array"
    VaseAssert.AssertArraysEqual TempArr, gNormalizedArray
    VaseAssert.AssertArraysEqual TempArr, gNonNormalArray
End Sub

Public Sub TestSize()
    VaseAssert.AssertEqual ArrayUtil.Size(gEmptyArray), 0, "Empty Array"
    VaseAssert.AssertEqual ArrayUtil.Size(gEmptyConstant), 0, "EMPTY"
    
    VaseAssert.AssertEqual ArrayUtil.Size(gNormalArray), 3, , "Normal Array"
    VaseAssert.AssertEqual ArrayUtil.Size(gNonNormalArray), 3, "Shifted Array"
End Sub

Public Sub TestCloneSize()
    VaseAssert.AssertArraysEqual ArrayUtil.CloneSize(gEmptyArray), Array(), "Empty Array"
    VaseAssert.AssertArraysEqual ArrayUtil.CloneSize(gEmptyConstant), Array(), "EMPTY"
    
    VaseAssert.AssertArraySize ArrayUtil.Size(gNormalArray), ArrayUtil.CloneSize(gNormalArray), "Normal Array"
    VaseAssert.AssertArraySize ArrayUtil.Size(gNonNormalArray), ArrayUtil.CloneSize(gNonNormalArray), "Shifted Array"
End Sub

Public Sub TestRemoveAllElements()
    Rewind_
    Setup
    VaseAssert.AssertArraysEqual ArrayUtil.RemoveAllElements("X", gEmptyArray), Array(), "Empty Array"
    VaseAssert.AssertArraysEqual ArrayUtil.RemoveAllElements("Y", gEmptyConstant), Array(), "EMPTY"

    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveAllElements(4, Array(1, 2, 3)), _
        Array(1, 2, 3)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveAllElements(2, Array(1, 2, 3)), _
        Array(1, 3)
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveAllElements(1, Array(1, "1", "One")), _
        Array("1", "One"), "Homegenous Array"
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveAllElements(1, Array(1, 1, 2)), _
        Array(2), "Multiple removal"
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveAllElements(1, Array(1, 1, 1)), _
        Array(), "All Removed"
        
    ArrayUtil.RemoveAllElements 1, Array(1, "1", Array(1))
    VaseAssert.AssertErrorNotRaised
    
    Ping_
End Sub

Public Sub TestIsInArr()
    VaseAssert.AssertFalse ArrayUtil.IsIn(0, gEmptyArray), "Empty Array"
    VaseAssert.AssertFalse ArrayUtil.IsIn(0, gEmptyConstant), "EMPTY"
    
    VaseAssert.AssertTrue ArrayUtil.IsIn(1, gNormalArray), "Normal Array"
    VaseAssert.AssertFalse ArrayUtil.IsIn(2, gNormalArray), "Normal Array"
End Sub

Public Sub TestRemoveDuplicates()
    VaseAssert.AssertEmptyArray ArrayUtil.RemoveDuplicates(gEmptyArray), "Empty Array"
    VaseAssert.AssertEmptyArray ArrayUtil.RemoveDuplicates(gEmptyConstant), "EMPTY"
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveDuplicates(Array(1, 2, 3, 2, 1)), Array(1, 2, 3)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.RemoveDuplicates( _
            Array("A", "b", "C", "C", "A")), _
            Array("A", "b", "C")
End Sub

Public Sub TestHasDuplicates()
    VaseAssert.AssertFalse ArrayUtil.HasDuplicates(gEmptyArray), "Empty Array"
    VaseAssert.AssertFalse ArrayUtil.HasDuplicates(gEmptyConstant), "EMPTY"
    
    VaseAssert.AssertTrue _
        ArrayUtil.HasDuplicates(Array(1, 2, 3, 2, 1))
    VaseAssert.AssertTrue _
        ArrayUtil.HasDuplicates(Array("A", "b", "C", "C", "A"))
End Sub

Public Sub TestJoinArrays()
    Dim LeftArr As Variant, RightArr As Variant
    LeftArr = Array(1, 2, 3)
    RightArr = Array(1, 2)

    VaseAssert.AssertEmptyArray ArrayUtil.JoinArrays(Empty, Array())
    VaseAssert.AssertEmptyArray ArrayUtil.JoinArrays(Array(), Empty)
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.JoinArrays(LeftArr, Empty), LeftArr
    VaseAssert.AssertArraysEqual _
        ArrayUtil.JoinArrays(Empty, RightArr), RightArr
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.JoinArrays(LeftArr, RightArr), _
        Array(1, 2, 3, 1, 2)
End Sub

Public Sub TestRange()
    VaseAssert.AssertEmptyArray _
        ArrayUtil.Range()
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(InclusiveRange:=True), _
        Array(0)
        
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Stop_:=10), _
        Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=5, Stop_:=10), _
        Array(5, 6, 7, 8, 9)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=5, Stop_:=10, Step_:=2), _
        Array(5, 7, 9)
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=0, Step_:=-1), _
        Array(10, 9, 8, 7, 6, 5, 4, 3, 2, 1)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=0, Step_:=-3), _
        Array(10, 7, 4, 1)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=1, Step_:=-3), _
        Array(10, 7, 4)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=2, Step_:=-3), _
        Array(10, 7, 4)
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=5, Stop_:=11, Step_:=2, InclusiveRange:=True), _
        Array(5, 7, 9, 11)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=5, Stop_:=11, Step_:=2, InclusiveRange:=False), _
        Array(5, 7, 9)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=1, Step_:=-3, InclusiveRange:=True), _
        Array(10, 7, 4, 1)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=1, Step_:=-3, InclusiveRange:=False), _
        Array(10, 7, 4)
        
    Ping_
End Sub

Public Sub TestCreateWithSize()
    VaseAssert.AssertEmptyArray _
        ArrayUtil.CreateWithSize(0)
    VaseAssert.AssertEmptyArray _
        ArrayUtil.CreateWithSize(-1)
        
    VaseAssert.AssertArraysEqual _
        ArrayUtil.CreateWithSize(1), Array(Empty)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.CreateWithSize(3), Array(Empty, Empty, Empty)
End Sub

Public Sub TestProjection()
    VaseAssert.AssertEmptyArray _
        ArrayUtil.Projection(Array(), Array())
    VaseAssert.AssertEmptyArray _
        ArrayUtil.Projection(Array(), Array(1))
    VaseAssert.AssertEmptyArray _
        ArrayUtil.Projection(Array(1), Array())

    Dim Arr As Variant, Indices As Variant
    Arr = ArrayUtil.Range(Start_:=1, Stop_:=25, Step_:=5)
    Indices = Array(3, 1)
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Projection(Indices, Arr), Array(16, 6)
End Sub

Public Sub TestSetDifference()
    VaseAssert.AssertEmptyArray _
        ArrayUtil.SetDifference(Array(), Array())
    VaseAssert.AssertEmptyArray _
        ArrayUtil.SetDifference(Array(), Array(1))
    VaseAssert.AssertEmptyArray _
        ArrayUtil.SetDifference(Array(1), Array())

    Dim Set_ As Variant, Subset_ As Variant, Elem_ As Variant, Diff_ As Variant
    Set_ = ArrayUtil.Range(Stop_:=5)
    Subset_ = Array(1, 3)
    Diff = ArrayUtil.SetDifference(Subset_, Set_)
    
    For Each Elem_ In Array(0, 2, 4)
        VaseAssert.AssertTrue _
            ArrayUtil.IsIn(Elem, Diff)
    Next
End Sub

Public Sub TestSlice()
    Dim Arr As Variant, Now_ As Date
    Now_ = Now
    Arr = Array(1, "A", True, Array(), Now_)

    VaseAssert.AssertEmptyArray _
        ArrayUtil.Slice(Array())
    
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Slice(Arr, 1, 3), _
        Array("A", True)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Slice(Arr, 0, 5, 2), _
        Array(1, True, Now_)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Slice(Arr, Start_:=4, Step_:=-2), _
        Array(Now_, True)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Slice(Arr, Start_:=4, Step_:=-2, InclusiveRange:=True), _
        Array(Now_, True, 1)
End Sub

Public Sub TestFirstAndLast()
    Dim Arr As Variant
    Arr = Array(1, 2, 3)
    
    VaseAssert.AssertEqual _
        ArrayUtil.First(Arr), 1
    VaseAssert.AssertEqual _
        ArrayUtil.Last(Arr), 3
        
    VaseAssert.AssertEqual _
        ArrayUtil.First(Array()), Empty
    VaseAssert.AssertEqual _
        ArrayUtil.Last(Array()), Empty
End Sub

Public Sub TestPartitionByIndices()
    Dim Arr As Variant, Parts As Variant
    Arr = ArrayUtil.Range(0, 30, 3)
    Parts = ArrayUtil.PartitionByIndices(Arr, Array(1, 3, 6))
    
    VaseAssert.AssertArraysEqual _
        Parts(0), Array(0)
    VaseAssert.AssertArraysEqual _
        Parts(1), Array(3, 6)
    VaseAssert.AssertArraysEqual _
        Parts(2), Array(9, 12, 15)
    VaseAssert.AssertArraysEqual _
        Parts(3), Array(18, 21, 24, 27)
End Sub

Public Sub TestIsAnyEmptyArray()
    VaseAssert.AssertFalse _
        ArrayUtil.IsAnyEmptyArray( _
            Array(True), Array(False), Array(0))
    VaseAssert.AssertTrue _
        ArrayUtil.IsAnyEmptyArray( _
            Array(True), Array(False), Array())
            
    VaseAssert.AssertFalse _
        ArrayUtil.IsAnyEmptyArray( _
            Array(True), Array(False))
    VaseAssert.AssertTrue _
        ArrayUtil.IsAnyEmptyArray( _
            Array(True), Array())
            
    VaseAssert.AssertFalse _
        ArrayUtil.IsAnyEmptyArray( _
            Array(True))
    VaseAssert.AssertTrue _
        ArrayUtil.IsAnyEmptyArray( _
            Array())
End Sub
