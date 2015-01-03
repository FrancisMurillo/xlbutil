Attribute VB_Name = "TestColUtil"
Public Function TestToArray()
    Dim Col As New Collection
    
    VaseAssert.AssertEmptyArray ColUtil.ToArray(Col)
        
    Col.Add "1"
    Col.Add 1
    Col.Add "Wroong"
    
    VaseAssert.AssertArraysEqual _
        ColUtil.ToArray(Col), Array("1", 1, "Wroong")
    
End Function
