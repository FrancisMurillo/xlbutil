Attribute VB_Name = "ChipInfo"
Public Sub WriteInfo()
    ChipReadInfo.References = Array() ' No dependencies
    ChipReadInfo.Modules = Array( _
        "ArrayUtil", _
        "AssertUtil", _
        "StringUtil")
End Sub


