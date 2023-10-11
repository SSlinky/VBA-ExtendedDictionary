VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright 2023 Sam Vanderslink
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy 
' of this software and associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights 
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
' copies of the Software, and to permit persons to whom the Software is 
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in 
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
' IN THE SOFTWARE.

Option Explicit

'-------------------------------------------------------------------------------
'   Class: TestResult
'   Simple data class containing a test result.
'-------------------------------------------------------------------------------

' Properties
'-------------------------------------------------------------------------------
Public Name As String
Public Failed As Boolean
Public Message As String

Public Property Get ToString() As String
Attribute ToString.VB_Description = "Returns a string representation of the class."
'   Returns a string representation of the class.
    Dim msg As String
    msg = Iif(Message = "", "", vbNewLine & "     " & Message)
    ToString = Iif(Failed, "!! Fail: ", "   Pass: ") & Name & msg
End Property


' Methods
'-------------------------------------------------------------------------------
Public Function AssertAreEqual( _
    expectedVal, _
    actualVal, _
    Optional forVal As String) As Boolean
Attribute AssertAreEqual.VB_Description = "Asserts expected is equal to actual."
'   Asserts expected is equal to actual.
'
'   Args:
'       expectedVal: The expected value.
'       actualVal: The actual value.
'       forVal: Additional context.
'
'   Returns:
'       False if the assertion passed.
'
    If expectedVal <> actualVal Then
        Failed = True
        SetExpectedActualMessage CStr(expectedVal), CStr(actualVal), forVal
    End If
    AssertAreEqual = Failed
End Function

Public Function AssertAreNotEqual( _
    expectedVal, _
    actualVal, _
    Optional forVal As String) As Boolean
Attribute AssertAreNotEqual.VB_Description = "Asserts values are not equal."
'   Asserts values are not equal.
'
'   Args:
'       expectedVal: The expected value.
'       actualVal: The actual value.
'       forVal: Additional context.
'
'   Returns:
'       False if the assertion passed.
'
    If expectedVal = actualVal Then
        Failed = True
        Message = "Inequality check [" & forVal & "]"
    End If
    AssertAreNotEqual = Failed
End Function

Public Function AssertIsTrue(val As Boolean, ctx As String) As Boolean
Attribute AssertIsTrue.VB_Description = "Asserts the value is true."
'   Asserts the value is true.
'
'   Args:
'       val: The value to test.
'       ctx: Context for the message.
'
'   Returns:
'       False if the assertion passed.
'
    If Not val Then
        Failed = True
        Message = "Failed " & ctx
    End If

    AssertIsTrue = Failed
End Function

Public Function AssertIsFalse(val As Boolean, ctx As String) As Boolean
Attribute AssertIsFalse.VB_Description = "Asserts the value is false."
'   Asserts the value is false.
'
'   Args:
'       val: The value to test.
'       ctx: Context for the message.
'
'   Returns:
'       False if the assertion passed.
'
    If val Then
        Failed = True
        Message = "Failed " & ctx
    End If

    AssertIsFalse = Failed
End Function

Public Function AssertNoException() As Boolean
Attribute AssertNoException.VB_Description = "Asserts no exception raised."
'   Asserts no exception raised.
'
'   Args:
'       errObj: The error object.
'
'   Returns:
'       False if the assertion passed.
'
    If Err <> 0 Then
        Failed = True
        Message = Err & " - " & Err.Description
    End If

    AssertNoException = Failed
End Function

Public Function AssertRaised(errNo As Long) As Boolean
Attribute AssertRaises.VB_Description = "Asserts a specific exception is raised."
'   Asserts a specific exception is raised.
'
'   Args:
'       errNo: The expected exception.
'
'   Returns:
'       False if the assertion passed.
'
    If Err = 0 Then
        Failed = True
        Message = "Expected exception " & errNo & " raised but none thrown."
    End If
    
    If Err <> errNo Then
        Failed = True
        Message = "Expected exception " & errNo & " raised but got " & Err & "."
    End If

    AssertRaised = Failed
End Function


' Helpers
'-------------------------------------------------------------------------------
Private Sub SetExpectedActualMessage(exp As String, act As String, Optional forVal As String)
Attribute SetExpectedActualMessage.VB_Description = "Sets a message based on expected vs actual values."
'   Sets a message based on expected vs actual values.
'
'   Args:
'       exp: The expected value.
'       act: The actual value.
'
    Dim msg As String
    msg = "Expected " & exp & " but got " & act
    Message = Iif(forVal = "", msg, msg & " [" & forVal & "]")
End Sub
