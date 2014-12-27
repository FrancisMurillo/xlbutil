Attribute VB_Name = "TestAssertUtil"
Public Sub TestArraysEqual()
    VaseAssert.AssertTrue AssertUtil.ArraysEqual(Empty, Array())
    VaseAssert.AssertTrue AssertUtil.ArraysEqual(Empty, Empty)
    VaseAssert.AssertTrue AssertUtil.ArraysEqual(Array(), Array())
    
    VaseAssert.AssertFalse AssertUtil.ArraysEqual(Array(1, 2), Array())
End Sub
