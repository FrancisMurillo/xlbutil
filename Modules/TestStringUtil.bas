Attribute VB_Name = "TestStringUtil"
Public Sub TestEditDistance()
    ' Some real computation would help
    VaseAssert.AssertEqual StringUtil.EditDistance("Ab", "Abc"), 67
    VaseAssert.AssertEqual StringUtil.EditDistance("Ab", "A"), 50
    
    ' Test string length
    Dim LongString As String
    LongString = "This is a very long string that the previous implementation of the Edit Distance would have failed to."
    
    VaseAssert.AssertEqual StringUtil.EditDistance(LongString, LongString), 100
    
    VaseAssert.Ping_
End Sub

Public Sub TestContains()
    VaseAssert.AssertTrue StringUtil.Contains("Abc", "A")
    VaseAssert.AssertTrue StringUtil.Contains("Abc", "Ab")
    VaseAssert.AssertFalse StringUtil.Contains("Abc", "D")
    VaseAssert.AssertFalse StringUtil.Contains("Abc", "Ac")
    
    VaseAssert.AssertFalse StringUtil.Contains("Abc", "a")
    VaseAssert.AssertTrue StringUtil.Contains("Abc", "a", IgnoreCase:=True)
    
    VaseAssert.Ping_
End Sub
