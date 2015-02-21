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
End Sub

Public Sub TestFindLike()
    Dim SArr As Variant
    SArr = Array("Hello", "Help", "Ouch", "Pouch")

    VaseAssert.AssertEqual _
        StringArrayUtil.FindLike("ASB", Array()), _
        -1
        
    VaseAssert.AssertEqual _
        StringArrayUtil.FindLike("Hel*", SArr), _
        0
    VaseAssert.AssertEqual _
        StringArrayUtil.FindLike("Hel*", SArr, StartIndex:=0 + 1), _
        1
    VaseAssert.AssertEqual _
        StringArrayUtil.FindLike("*ouch", SArr), _
        3
    VaseAssert.AssertEqual _
        StringArrayUtil.FindLike("*ouch", SArr, IgnoreCase:=True), _
        2
End Sub

