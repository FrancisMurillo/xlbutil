Attribute VB_Name = "TestStringArrayUtil"
Public Sub TestTrimAll()
    VaseAssert.AssertEmptyArray _
        StringArrayUtil.TrimAll(Array())
    
    VaseAssert.AssertArraysEqual _
        StringArrayUtil.TrimAll( _
            Array("", " Left Strip", "Right Strip  ", "   Strip    ")), _
            Array("", "Left Strip", "Right Strip", "Strip")
End Sub

Public Sub TestIsInLike()
    Dim SArr As Variant
    SArr = Array("Hello", "Help", "Ouch", "Pouch")

    VaseAssert.AssertFalse _
        StringArrayUtil.IsInLike("ASB", Array())
    
        
    VaseAssert.AssertTrue _
        StringArrayUtil.IsInLike("Hel*", SArr)
    VaseAssert.AssertFalse _
        StringArrayUtil.IsInLike("HeL*", SArr)
    VaseAssert.AssertTrue _
        StringArrayUtil.IsInLike("HeL*", SArr, IgnoreCase:=True)
        
    VaseAssert.AssertTrue _
        StringArrayUtil.IsInLike("*uch", SArr)
    VaseAssert.AssertTrue _
        StringArrayUtil.IsInLike("*uc*", SArr)
    
    Ping_
End Sub
