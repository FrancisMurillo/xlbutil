Attribute VB_Name = "TestFnArrayUtil"
Public Sub TestMap()
    Dim NumArr As Variant, StrArr As Variant, VarArr As Variant
    NumArr = Array(1, 2, 3, 2, 1)
    StrArr = Array("I", "Me", "Mine")
    VarArr = Array(1, "2", True, Empty)
    
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.Map("", Array())
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map("FnTestLambda.Negative_", NumArr), _
        Array(-1, -2, -3, -2, -1)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map("FnTestLambda.Prefix_", StrArr), _
        Array("Pre: I", "Pre: Me", "Pre: Mine")
    
    Dim ActVarArr As Variant, Pair As Variant
    ActVarArr = Map("FnTestLambda.WrapArray_", VarArr)
    

    Ping_
End Sub
