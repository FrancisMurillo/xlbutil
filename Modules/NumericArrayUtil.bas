Attribute VB_Name = "NumericArrayUtil"
' Numeric Array Utility
' ---------------------
'
' These utilities like StringArrayUtil are for numeric arrays

'# Sums the numbers in an array
'! This is not typed, so if the array contains Double, Decimal or Long;
'! this might produced different types
Public Function Sum(NArr As Variant) As Variant
    Sum = 0
    If ArrayUtil.IsEmptyArray(NArr) Then _
        Exit Function

    Dim Val As Variant
    For Each Val In NArr
        Sum = Sum + Val
    Next
End Function

'# Just like Sum, this takes the product
Public Function Product(NArr As Variant) As Variant
    Product = 1
    If ArrayUtil.IsEmptyArray(NArr) Then _
        Exit Function

    Dim Val As Variant
    For Each Val In NArr
        Product = Product * Val
    Next
End Function

