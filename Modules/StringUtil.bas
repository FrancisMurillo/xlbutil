Attribute VB_Name = "StringUtil"
'===========================
'--- Module Contract     ---
'===========================
' There is no contract for Strings, aside for type safety


' Checks if either strings starts with the other
' Used in the truncated case of data compare
Public Function startsEither(str As String, Prefix As String) As Boolean
    startsEither = startsWith(str, Prefix) Or startsWith(Prefix, str)
End Function

' Reference: http://stackoverflow.com/questions/20802870/vba-test-if-string-begins-with-a-string
Public Function startsWith(str As String, Prefix As String) As Boolean
    startsWith = Left(str, Len(Prefix)) = Prefix
End Function

' String Containment
Public Function ContainsBoth(Left As String, Right As String) As Boolean
    ContainsBoth = Contains(Left, Right) Or Contains(Right, Left)
End Function


'# String Contains
'# Note the argument interchange, feels more functional
Public Function Contains(Source As String, Target As String, Optional IgnoreCase As Boolean = False) As Boolean
    If Not IgnoreCase Then
        Contains = (InStr(Source, Target) > 0)
    Else
        Contains = (InStr(UCase(Source), UCase(Target)) > 0)
    End If
End Function

'# Same as Contains(), except the arguments are interchanged
'# It feels more functional to have the target before the source
Public Function In_(Target As String, Source As String, Optional IgnoreCase As Boolean = False) As Boolean
    In_ = Contains(Source, Target)
End Function






' Returns a string good enough to match for headers
' Checks for substring then edit distance
' An abstraction for the selection criteria, could be even more complicated
Public Function GoodEnoughMatch(Matchee As String, Choices As Variant) As String
    
    ' Check edit distance
    GoodEnoughMatch = BestMatch(Matchee, Choices, 75)
    
    If GoodEnoughMatch <> Empty Then
        Exit Function
    End If
    ' Checks in string
    Dim Choice As Variant
    For Each Choice In Choices
        If (Choice <> "" And Matchee <> "") And (InStr(Choice, Matchee) > 0 Or InStr(Matchee, Choice) > 0) Then
            GoodEnoughMatch = Choice
            Exit Function
        End If
    Next
End Function

' Gets the best match of an string from a list of strings.
' It returns a string in the list if the rating is above or equal the threshold, else Empty or ""
' Note: If there are multiple best matches, there is no guarantee which one is chosen
Public Function BestMatch(Matched As String, Choices As Variant, Optional Threshold As Integer = 0) As String
    Dim Choice As Variant, CurRating As Integer
    Dim BestRating As Integer
    CurRating = 0
    BestMatch = Empty
    BestRating = Threshold - 1
    For Each Choice In Choices
        CurRating = Levenshtein3(Matched, Choice)
        If BestRating < CurRating Then
            BestRating = CurRating
            BestMatch = Choice
        End If
    Next
End Function
Public Function BestMatchIndex(Matched As String, Choices() As String, Optional Threshold As Integer = 0) As String
    Dim BestRating As Integer
    Dim CurRating As Integer, Index As Integer
    CurRating = 0
    BestMatchIndex = -1
    For Index = LBound(Choices) To UBound(Choices)
        CurRating = Levenshtein3(Matched, Choices(Index))
        If BestRating < CurRating Then
            BestRating = CurRating
            BestMatchIndex = Index
        End If
    Next
End Function

'# Computes the simple edit distance between two strings ranging from 1 to 100
'? Reference: http://stackoverflow.com/questions/4243036/levenshtein-distance-in-excel
Public Function EditDistance(ByVal string1 As String, ByVal string2 As String) As Long
    Dim I As Long, J As Long, string1_length As Long, string2_length As Long
    Dim distance() As Long, smStr1() As Long, smStr2() As Long
    Dim min1 As Long, min2 As Long, min3 As Long, minmin As Long, MaxL As Long
    
    string1_length = Len(string1):  string2_length = Len(string2)
    If string1_length = 0 Then
        EditDistance = 0
        Exit Function
    End If
    If string2_length = 0 Then
        EditDistance = 0
        Exit Function
    End If
    
    ' Resize arrays, modified constants to adjust to variable string lengths plus optimization
    ReDim distance(0 To string1_length, 0 To string2_length)
    ReDim smStr1(1 To string1_length)
    ReDim smStr2(1 To string2_length)
    
    distance(0, 0) = 0
    For I = 1 To string1_length:    distance(I, 0) = I: smStr1(I) = Asc(LCase(Mid$(string1, I, 1))): Next
    For J = 1 To string2_length:    distance(0, J) = J: smStr2(J) = Asc(LCase(Mid$(string2, J, 1))): Next
    For I = 1 To string1_length
        For J = 1 To string2_length
            If smStr1(I) = smStr2(J) Then
                distance(I, J) = distance(I - 1, J - 1)
            Else
                min1 = distance(I - 1, J) + 1
                min2 = distance(I, J - 1) + 1
                min3 = distance(I - 1, J - 1) + 1
                If min2 < min1 Then
                    If min2 < min3 Then minmin = min2 Else minmin = min3
                Else
                    If min1 < min3 Then minmin = min1 Else minmin = min3
                End If
                distance(I, J) = minmin
            End If
        Next
    Next
    
    ' Levenshtein3 will properly return a percent match (100%=exact) based on similarities and Lengths etc...
    MaxL = string1_length: If string2_length > MaxL Then MaxL = string2_length
    EditDistance = 100 - CLng((distance(string1_length, string2_length) * 100) / MaxL)
End Function



