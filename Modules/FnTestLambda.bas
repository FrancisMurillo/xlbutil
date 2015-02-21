Attribute VB_Name = "FnTestLambda"
' Unit Test Functions
' -------------------
'
' These function are used in the unit testing, not used in production
' As stated, these functions set the result variable instead

'# Turns a number to its negative
Public Sub Negative_(Val As Long)
    FnLambda.Result = -1 * Val
End Sub

'# Adds a prefix to a string, 'Pre: ' prefix
Public Sub Prefix_(Val As String)
    FnLambda.Result = "Pre: " & Val
End Sub

'# Just wraps the value into an array
Public Sub WrapArray_(Val As Variant)
    FnLambda.Result = Array(Val)
End Sub

'# Accepts only 2
Public Sub IsTwo_(Val As Long)
    FnLambda.Result = (Val = 2)
End Sub

'# Accepts only Francis
Public Sub IsFrancis_(Val As String)
    FnLambda.Result = (Val = "Francis")
End Sub

'# Accepts all, a default filter for All
Public Sub True_(Val As Variant)
    FnLambda.Result = True
End Sub
