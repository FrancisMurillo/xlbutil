Attribute VB_Name = "TestFn"
Public Sub TestInvoke()
    VaseAssert.AssertEqual _
        Fn.Invoke("FnFunction.Identity_", Array(1)), _
        1
End Sub

Public Sub TestCurry()
    Dim AddFive As String, AddFiveAndFour As String, AddNine As String, Add_L As String
    Dim ConstOne As String
    AddFive = Fn.Curry("FnOperator.Add_", Array(5))
    
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFive, Array(4)), _
        (5 + 4)
        
    AddFiveAndFour = Fn.Curry(AddFive, Array(4))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFour, Array()), _
        9
        
    AddNine = Fn.Curry("FnOperator.Add_", Array(4, 5))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFour, Array()), _
        Fn.Invoke(AddNine, Array())
        
    Add_L = Fn.Curry("FnOperator.Add_", Array())
    VaseAssert.AssertEqual _
        Fn.Invoke(Add_L, Array(4, 5)), _
        9
        
    ConstOne = Fn.Curry("FnFunction.Identity_", Array(1))
    VaseAssert.AssertEqual _
        Fn.Invoke(ConstOne, Array()), _
        1
End Sub

Public Sub TestCompose()
    Dim NegRecFn As String
    
    NegRecFn = Fn.Compose(Array("FnFunction.Negative_", "FnFunction.Reciprocal_"))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(NegRecFn, 2), _
        -(1 / 2)
       
    Dim RemoveAandI_Fn As String, ToUpperAndRemove_Fn As String
    
    RemoveAandI_Fn = Fn.Compose(Array("FnTestLambda.RemoveA_", "FnTestLambda.RemoveI_"))
    ToUpperAndRemove_Fn = Fn.Compose(Array("FnTestLambda.ToUppercase_", RemoveAandI_Fn))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(ToUpperAndRemove_Fn, "Francis"), _
        "FRNCS"
End Sub

Public Sub TestReinvoke()
    Dim WithTwoAndThree As String
    WithTwoAndThree = Fn.Reinvoke(Array(2, 3))
    
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThree, "FnOperator.Add_"), _
        5
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThree, "FnOperator.Multiply_"), _
        6
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(WithTwoAndThree, _
            Array("FnOperator.Add_", "FnOperator.Multiply_")), _
        Array(5, 6)
    
    Ping_
End Sub

