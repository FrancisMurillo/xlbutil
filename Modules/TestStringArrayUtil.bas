Attribute VB_Name = "TestStringArrayUtil"
Public Sub TestTrimAll()
    VaseAssert.AssertEmptyArray _
        StringArrayUtil.TrimAll(Array())
    
    VaseAssert.AssertArraysEqual _
        StringArrayUtil.TrimAll( _
            Array("", " Left Strip", "Right Strip  ", "   Strip    ")), _
            Array("", "Left Strip", "Right Strip", "Strip")
End Sub
