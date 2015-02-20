Attribute VB_Name = "FnLambda"
' Functional Lambdas
' ------------------
'
' This module contains functions to be used as functions for FnArrayUtil.
'
' Functions here are declared as sub or routines instead since these are called by Application.Run,
' so to pass the output the return is set to the module variable Result and FnArray methods will use that instead.
' Such a sad way to make lambdas but better than nothing
'
' Moving on, place the functions here to segregate them from normally functioning methods.
' Put a suffix _ at the end of the functions as a convention telling them they are fitted functions for this module
'
' However, there is no strict requirement to put all the functions here and follow the convention;
' the only thing required is to set the variable result
'
' I do lament the step in wrapping functions to module and fitting the return to the variable.
' Anyway, I do suggest you put the wrapped functions here as not to affect the writing to other modules

' The Result variable, place your result here
Public Result As Variant


