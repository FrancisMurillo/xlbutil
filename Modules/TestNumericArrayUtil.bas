Attribute VB_Name = "TestNumericArrayUtil"
Public Sub TestSum()
    VaseAssert.AssertEqual _
        NumericArrayUtil.Sum(Array()), 0
        
    VaseAssert.AssertEqual _
        NumericArrayUtil.Sum(Array(1#)), 1
    VaseAssert.AssertEqual _
        NumericArrayUtil.Sum(Array(1#, CLng(2))), 3
    VaseAssert.AssertEqual _
        NumericArrayUtil.Sum(Array(1#, CLng(2), CDec(3))), 6
End Sub

Public Sub TestProduct()
    VaseAssert.AssertEqual _
        NumericArrayUtil.Product(Array()), 1
        
    VaseAssert.AssertEqual _
        NumericArrayUtil.Product(Array(1#)), 1
    VaseAssert.AssertEqual _
        NumericArrayUtil.Product(Array(1#, CLng(2))), 2
    VaseAssert.AssertEqual _
        NumericArrayUtil.Product(Array(1#, CLng(2), CDec(3))), 6
End Sub

