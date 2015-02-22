Attribute VB_Name = "PathUtil"
' File Path Utilities
' -------------------
'
' Utilities to manage the file path, taken from Java and Python

'# Joins paths, the Pythonic way of os.path
Public Function Join_(BasePath As String, ParamArray ExtPaths() As Variant)
    Dim Path_ As String, ExtPath As Variant, ExtPath_ As String
    
    Path_ = RemoveOptionalPathSeparator(BasePath)
    
    For Each ExtPath In ExtPaths
        ExtPath_ = RemoveOptionalPathSeparator(CStr(ExtPath))
        If ExtPath <> "" Then
            Path_ = Path_ & Application.PathSeparator & ExtPath_
        End If
    Next
    
    Join_ = Path_
End Function

Private Function RemoveOptionalPathSeparator(Path As String) As String
    Dim Path_ As String
    Path_ = Path
    If Right(Path_, 1) = Application.PathSeparator Then
        Path_ = Left(Path_, Len(Path_) - 1)
    End If
    If Left(Path_, 1) = Application.PathSeparator Then
        Path_ = Right(Path_, Len(Path_) - 1)
    End If
    RemoveOptionalPathSeparator = Path_
End Function
