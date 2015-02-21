Attribute VB_Name = "TestFnArrayUtil"
Public Sub TestZip()
    Dim Arr As Variant
    
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.Zip(Array( _
            ArrayUtil.Range(0, 1), _
            ArrayUtil.Range(0, 2), _
            Array()))
    
    Dim ActArr As Variant
    ActArr = FnArrayUtil.Zip(Array( _
                ArrayUtil.Range(0, 5, 3), _
                ArrayUtil.Range(-10, 10, 7)))
    VaseAssert.AssertArraySize 2, ActArr
    VaseAssert.AssertArraysEqual _
        ActArr(0), _
        Array(0, -10)
    VaseAssert.AssertArraysEqual _
        ActArr(1), _
        Array(3, -3)
        
    Ping_
End Sub

Public Sub TestChain()
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.Chain( _
            Array(), Array(), Array())
    
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Chain( _
            ArrayUtil.Range(0, 4, 2), _
            ArrayUtil.Range(4, 8, 2), _
            ArrayUtil.Range(8, 12, 2)), _
        ArrayUtil.Range(0, 12, 2)
End Sub

Public Sub TestTakeN()
    Dim Arr As Variant
    Arr = Array(1, "A", Empty)
    
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.TakeN(0, Array())
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.TakeN(1, Array())
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.TakeN(-1, Array(1, 2, 3))
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.TakeN(1, Arr), Array(1)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.TakeN(2, Arr), Array(1, "A")
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.TakeN(3, Arr), Array(1, "A", Empty)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.TakeN(4, Arr), Array(1, "A", Empty)
End Sub


Public Sub TestZipWith()
    Dim Arr As Variant
    
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.ZipWith_("FnFunction.Identity_", Array( _
            ArrayUtil.Range(0, 1), _
            ArrayUtil.Range(0, 2), _
            Array()))
    
    Dim ActArr As Variant, ZipActArr As Variant
    ActArr = FnArrayUtil.ZipWith_("FnFunction.Identity_", Array( _
                ArrayUtil.Range(0, 5, 3), _
                ArrayUtil.Range(-10, 10, 7)))
    VaseAssert.AssertArraySize 2, ActArr
    VaseAssert.AssertArraysEqual _
        ActArr(0), _
        Array(0, -10)
    VaseAssert.AssertArraysEqual _
        ActArr(1), _
        Array(3, -3)
        
    ZipActArr = FnArrayUtil.ZipWith_("FnTestLambda.Formula_", Array( _
                ArrayUtil.Range(0, 5, 3), _
                ArrayUtil.Range(-10, 10, 7)))
    VaseAssert.AssertArraysEqual _
        ZipActArr, Array(0, -9)
End Sub

