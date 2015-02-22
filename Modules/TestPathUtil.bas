Attribute VB_Name = "TestPathUtil"
Public Sub TestJoin()
    Dim BasePath As String
    BasePath = "C:\usr\"
    
    VaseAssert.AssertEqual _
        PathUtil.Join_(BasePath, ""), "C:\usr"
        
    VaseAssert.AssertEqual _
        PathUtil.Join_(BasePath, "\meow"), "C:\usr\meow"
    VaseAssert.AssertEqual _
        PathUtil.Join_(BasePath, "\meow", "roar\"), "C:\usr\meow\roar"
    VaseAssert.AssertEqual _
        PathUtil.Join_(BasePath, "\meow", "roar\", "\rawr\"), "C:\usr\meow\roar\rawr"
    
    VaseAssert.AssertEqual _
        PathUtil.Join_("\relative", ""), "relative"
    VaseAssert.AssertEqual _
        PathUtil.Join_("\relative", "path\"), "relative\path"
End Sub
