Attribute VB_Name = "TestFnArrayUtil"
Public Sub TestMap_()
    Dim NumArr As Variant, StrArr As Variant, VarArr As Variant
    NumArr = Array(1, 2, 3, 2, 1)
    StrArr = Array("I", "Me", "Mine")
    VarArr = Array(1, "2", True, Empty)
    
    VaseAssert.AssertEmptyArray _
        FnArrayUtil.Map_("", Array())
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_("FnTestLambda.Negative_", NumArr), _
        Array(-1, -2, -3, -2, -1)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_("FnTestLambda.Prefix_", StrArr), _
        Array("Pre: I", "Pre: Me", "Pre: Mine")
    
    Dim ActVarArr As Variant, Pair As Variant
    ActVarArr = Map_("FnTestLambda.WrapArray_", VarArr)
    
    For Each Pair In ArrayUtil.Zip(ActVarArr, VarArr)
        VaseAssert.AssertEqual _
            Pair(0)(0), Pair(1)
    Next
    
End Sub

Public Function TestFilter_()
    Dim NumArr As Variant, StrArr As Variant, VarArr As Variant
    NumArr = Array(1, 2, 3, 2, 1)
    StrArr = Array("I", "Me", "Mine")
    VarArr = Array(1, "2", True, Empty)

    VaseAssert.AssertEmptyArray _
        FnArrayUtil.Filter_("", Array())

    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Filter_("FnTestLambda.IsTwo_", NumArr), _
        Array(2, 2)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Filter_("FnTestLambda.IsFrancis_", StrArr), _
        Array()
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Filter_("FnTestLambda.True_", VarArr), _
        VarArr
End Function
