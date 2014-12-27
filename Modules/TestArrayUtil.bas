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
    
    TempArr = ArrayUtil.ShiftBase(gNormalizedArray)
    VaseAssert.AssertEqual LBound(TempArr), 0, "Normalized array"
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
