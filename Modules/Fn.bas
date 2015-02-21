Attribute VB_Name = "Fn"
' Functional Programming: Fn
' ------------------
'
' This module provides a mechanism to provide pseudo lambda in VBA.
'
' # Module Definition
'
' Since we don't have first class functions, their names could be their pointers instead.
' By using Application.Run, we can invoke the function.
' Sadly, Application.Run has some limitation like not having a return value;
' This is remedied by setting the property Result(should have been Return but its already a keyword)
' with the value of the function.
'
' So if you have a function in a module, MyModule, like so...
'
' Public Function MyFunc(MyArg as Variant) As String
'   MyFunc = Str(MyArg)
' End Function
'
' Under the definition of pseudo lambda it will be...
'
' Public Sub MyFunc(MyArg as Variant)
'   Fn.Result = Str(MyArg)
' End Sub
'
' This pseudo function can be invoked by...
'
' Fn.Invoke("MyModule.MyFunc", Array(MyArg))
'
' Not a whole lot of difference except for Function being Sub
' and the return mechanism, which I say might be better than writing the function name all the time,
' and the invokation mechanism, something of a necessary evil
'
' So what does wrapping or fitting the function to lambdas get you?
' It gives you the ability to invoke them with cool functional methods like Map, Reduce and Filter.
' If you tasted the first kiss of functional programming, this is a little drop of heaven.
' This is better than doing procedural and crappy OO.
'
' Modules with the Fn prefix(except this one) utilize this mechanism. If you are making your own, you should too for convention.
' Methods utilizing the Fn.Invoke should be placed in the modules with Fn and end with undersocre for convention.
' Not a requirement, but helps with familiarity although I use this convention as well to avoid naming conflicts.
'
' Word of warning, this mechanism trades flexibility for performance.
' So when using this for performance critical aspects, take of your gloves and get your hands dirty.
'
' So be mindful when and where to use this. Such is the way of the programmer.

' ## Consants
'
' Error constants
Private Const ERR_SOURCE As String = "Fn"
Private Const ERR_OFFSET As Long = 2000

' ## Property
'
' The Result property, place your result here. Write-only, that's what it's supposed to be.
Private gResult As Variant
Public Property Let Result(Val As Variant)
    gResult = Val
End Property

' ## Functions

'# Invokes a function given its name and an array of arguments
'# This is achieved by using Application.Run and the concept of functions just have one argument
'# There is one subtle limitation, the maximum number of arguments. Due to Application.Run, the maximum number is 30.
'# Anything higher would result in an error
'P MethodName: The method to be invoked given its name. It should be the full name to be exact like "Fn.Invoke", not just "Invoke".
'P             You can just use the method name but you might run a function of the same name, so long name for safety.
'P Args: This is the arguments for the method wrapped in an array.
'P       This is also assumed to have base zero, but not a strict condition.
'P       The arguments are applied by order not by index, but make my our easier by using Array() to wrap the arguments
Public Function Invoke(MethodName As String, Args As Variant) As Variant
On Error GoTo ErrHandler:
    Dim Size_ As Long, Args_ As Variant
    Args_ = ArrayUtil.AsNormalArray(Args)
    Size_ = ArrayUtil.Size(Args_)
    
    gResult = Empty
    
    ' The long case of Application.Run, Python FTW
    Select Case Size_
        Case 0
            Application.Run MethodName
        Case 1
            Application.Run MethodName, Args_(0)
        Case 2
            Application.Run MethodName, Args_(0), Args_(1)
        Case 3
            Application.Run MethodName, Args_(0), Args_(1), Args_(2)
        Case 4
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3)
        Case 5
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4)
        Case 6
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5)
        Case 7
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6)
        Case 8
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7)
        Case 9
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8)
        Case 10
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9)
        Case 11
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10)
        Case 12
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11)
        Case 13
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12)
        Case 14
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13)
        Case 15
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14)
        Case 16
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15)
        Case 17
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16)
        Case 18
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17)
        Case 19
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18)
        Case 20
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19)
        Case 21
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20)
        Case 22
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21)
        Case 23
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22)
        Case 24
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23)
        Case 25
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24)
        Case 26
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24), Args_(25)
        Case 27
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24), Args_(25), Args_(26)
        Case 28
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24), Args_(25), Args_(26), Args_(27)
        Case 29
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24), Args_(25), Args_(26), Args_(27), Args_(28)
        Case 30
            Application.Run MethodName, Args_(0), Args_(1), Args_(2), Args_(3), Args_(4), Args_(5), Args_(6), Args_(7), Args_(8), Args_(9), Args_(10), Args_(11), Args_(12), Args_(13), Args_(14), Args_(15), Args_(16), Args_(17), Args_(18), Args_(19), Args_(20), Args_(21), Args_(22), Args_(23), Args_(24), Args_(25), Args_(26), Args_(27), Args_(28), Args_(29)
        Case Else
            Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Invoking " & MethodName & " with " & Size_ & " arguments exceeded the maximum number(30)"
    End Select
    
    Invoke = gResult
ErrHandler:
    If Err.Number = 1004 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "The method " & MethodName & " does not exist"
    ElseIf Err.Number <> 0 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, MethodName & " caused an error: " & Err.Description
    End If
End Function