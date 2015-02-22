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

