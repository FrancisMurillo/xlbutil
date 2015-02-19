Attribute VB_Name = "TestTupleUtil"
Public Sub TestTranspose()
    Dim Tuples As Variant, TTuples As Variant
    Tuples = Array(Array(1, 2, 3), Array(2, 3, 4))
    
    TTuples = TupleUtil.Transpose(Tuples)
    

End Sub
