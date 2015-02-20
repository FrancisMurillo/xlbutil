Attribute VB_Name = "AppUtil"
Public Function IsUppercase(Phrase As String)
    IsUppercase = (UCase(Phrase) = Phrase)
End Function

Public Function SheetLastRow(Sheet As Worksheet) As Integer
    SheetLastRow = Sheet.Range("A" & Sheet.Rows.Count).End(xlUp).Row
End Function

Public Function IsTitlePosition(Phrase As String)
    IsTitlePosition = (Right(Phrase, 1) = ":")
End Function

Public Function TrimStringArray(SArr As Variant)
    If ArrayUtil.IsEmptyArray(SArr) Then
        TrimStringArray = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If
    
    Dim Arr_ As Variant, Index As Long
    Arr_ = ArrayUtil.CloneSize(SArr)
    For Index = LBound(SArr) To UBound(SArr)
        Arr_(Index) = Trim(SArr(Index))
    Next
    
    TrimStringArray = Arr_
End Function

Public Function SpliceArray(Arr As Variant, Optional Start_ As Long = 0, Optional Stop_ As Long = 0, Optional Step_ As Long = 1)
    If ArrayUtil.IsEmptyArray(Arr) Then
        SpliceArray = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Rng As Variant
    Rng = ArrayUtil.Range(Start_, Stop_, Step_)

    SpliceArray = ArrayUtil.Projection(Rng, Arr)
End Function
