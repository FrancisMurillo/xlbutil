Attribute VB_Name = "TestBooleanArrayUtil"
Public Sub TestAny()
    VaseAssert.AssertFalse _
        BooleanArrayUtil.Any_( _
            Array())
    
    VaseAssert.AssertTrue _
        BooleanArrayUtil.Any_( _
            Array(True))
    VaseAssert.AssertFalse _
        BooleanArrayUtil.Any_( _
            Array(False))
            
    VaseAssert.AssertTrue _
        BooleanArrayUtil.Any_( _
            Array(True, False))
    VaseAssert.AssertTrue _
        BooleanArrayUtil.Any_( _
            Array(False, True))
            
    VaseAssert.AssertTrue _
        BooleanArrayUtil.Any_( _
            Array(True, True, True))
    VaseAssert.AssertFalse _
        BooleanArrayUtil.Any_( _
            Array(False, False, False))
End Sub

Public Sub TestAll()
    VaseAssert.AssertFalse _
        BooleanArrayUtil.All_( _
            Array())
    
    VaseAssert.AssertTrue _
        BooleanArrayUtil.All_( _
            Array(True))
    VaseAssert.AssertFalse _
        BooleanArrayUtil.All_( _
            Array(False))
            
    VaseAssert.AssertFalse _
        BooleanArrayUtil.All_( _
            Array(True, False))
    VaseAssert.AssertFalse _
        BooleanArrayUtil.All_( _
            Array(False, True))
            
    VaseAssert.AssertTrue _
        BooleanArrayUtil.All_( _
            Array(True, True, True))
    VaseAssert.AssertFalse _
        BooleanArrayUtil.All_( _
            Array(False, False, False))
End Sub

