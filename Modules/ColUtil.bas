Attribute VB_Name = "ColUtil"
'# Converts an collection to an array
Public Function ToArray(Col As Variant)
    If Col.Count = 0 Then
        ToArray = Array()
    Else
        Dim Index As Long, Arr_ As Variant, Item As Variant
        Arr_ = ArrayUtil.CreateWithSize(Col.Count)
        Index = 0
        For Each Item In Col
            If IsObject(Item) Then
                Set Arr_(Index) = Item
            Else
                Arr_(Index) = Item
            End If
            Index = Index + 1
        Next
        
        ToArray = Arr_
    End If
End Function
