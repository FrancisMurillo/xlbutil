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
    VaseAssert.AssertFalse ArrayUtil.IsInArray(0, gEmptyArray), "Empty Array"
    VaseAssert.AssertFalse ArrayUtil.IsInArray(0, gEmptyConstant), "EMPTY"
    
    VaseAssert.AssertTrue ArrayUtil.IsInArray(1, gNormalArray), "Normal Array"
    VaseAssert.AssertFalse ArrayUtil.IsInArray(2, gNormalArray), "Normal Array"
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
        Array(10, 7, 4, 1)
    VaseAssert.AssertArraysEqual _
        ArrayUtil.Range(Start_:=10, Stop_:=2, Step_:=-3), _
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