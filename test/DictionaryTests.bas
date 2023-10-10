Attribute VB_Name = "DictionaryTests"
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


Private passTests As New Collection
Private failTests As New Collection

Public Sub RunTests()
Attribute RunTests.VB_Description = "Runs all tests."
'   Runs all tests.
'
    Set passTests = New Collection
    Set failTests = New Collection

    Dim testName As Variant
    For Each testName In GetTestNames()
        RunTest CStr(testName)
    Next testName

    Dim p As Long, f As Long
    p = passTests.Count
    f = failTests.Count

    Debug.Print "-------------------------------------------"
    Debug.Print " Passed: " & p & " (" & Format(p / (p + f), "0.00%)")
    Debug.Print " Failed: " & f & " (" & Format(f / (p + f), "0.00%)")
    Debug.Print "-------------------------------------------"
    
End Sub

Private Sub RunTest(testName As String)
Attribute RunTest.VB_Description = "Runs the named test and stores the result."
'   Runs the named test and stores the result.
'
'   Args:
'       testName: The name of the function returning a TestResult.
'
    Dim tr As TestResult
    Set tr = Application.Run(testName)
    tr.Name = testName
    
    If tr.Failed Then
        failTests.Add tr
        Debug.Print "FAIL: " & tr.Name, tr.Message
    Else
        passTests.Add tr
    End If
End Sub

Private Function GetTestNames() As Collection
Attribute GetTestNames.VB_Description = "Gets the test names from this module."
'   Gets the test names from this module.
'   A valid test starts with Private Function TestDictionary_ and takes no args.
'
'   Returns:
'       A collection of strings representing names of tests.
'
    Const MODULENAME As String = "DictionaryTests"
    Const FUNCTIONID As String = "Private Function "
    Const TESTSTARTW As String = "Private Function TestDictionary_"

    Dim tswLen As Long
    tswLen = Len(TESTSTARTW)

    Dim codeMod As Object
    Set codeMod = ThisWorkbook.VBProject.VBComponents(MODULENAME).CodeModule

    Dim i As Long
    Dim results As New Collection
    For i = 1 To codeMod.CountOfLines
        Dim lineContent As String
        lineContent = codeMod.Lines(i, 1)

        If Left(lineContent, tswLen) = TESTSTARTW Then
            Dim funcName As String
            funcName = Split(Split(lineContent, FUNCTIONID)(1), "(")(0)
            results.Add funcName
        End If
    Next i
    
    Set GetTestNames = results
End Function

Private Function TestDictionary_Add() As TestResult
Attribute TestDictionary_Add.VB_Description = "Add an item to the dictionary."
'   Add an item to the dictionary.
    Dim tr As New TestResult

'   Arrange
    Const ADDKEY As String = "K"
    Const ADDVAL As String = "V"
    Dim d As New Dictionary

'   Act
    On Error Resume Next
    d.Add ADDKEY, ADDVAL

'   Assert
    If Err <> 0 Then
        tr.Failed = True
        tr.Message = "Exception unexpected: " _
            & Err & " - " & Err.Description
        GoTo Finally
    End If
    On Error GoTo 0

    If d.Count <> 1 Then
        tr.Failed = True
        tr.Message = "Expected count 1 but got " & d.Count
        GoTo Finally
    End If

    If d(ADDKEY) <> ADDVAL Then
        tr.Failed = True
        tr.Message = "Expected count 1 but got " & d.Count
        GoTo Finally
    End If

Finally:
    Set TestDictionary_Add = tr
End Function

Private Function TestDictionary_AddBulkColMode() As TestResult
Attribute TestDictionary_AddBulkColMode.VB_Description = "Add bulk items to the dictionary."
'   Add bulk items to the dictionary.
    Dim tr As New TestResult

'   Arrange
    Dim bulkData() As Variant
    ReDim bulkData(1 To 3, 1 To 4)

    Const HDRS As String = "ABC"
    Const VALS As String = " 123"

    Dim i As Long
    For i = 1 To UBound(bulkData, 1)
        Dim j As Long
        For j = 1 To UBound(bulkData, 2)
            If j = 1 Then
                bulkData(i, j) = Mid(HDRS, i, 1)
            Else
                bulkData(i, j) = Mid(HDRS, i, 1) & Mid(VALS, j, 1)
            End If
        Next j
    Next i

    Dim d As New Dictionary

'   Act
    On Error Resume Next
    d.AddBulk bulkData

'   Assert
    If Err <> 0 Then
        tr.Failed = True
        tr.Message = "Exception unexpected: " _
            & Err & " - " & Err.Description
        GoTo Finally
    End If
    On Error GoTo 0

    If d.Count <> Len(HDRS) Then
        tr.Failed = True
        tr.Message = "Expected count " & HDRS & " but got " & d.Count
        GoTo Finally
    End If

    For i = 1 To UBound(bulkData, 1)
        For j = 2 To UBound(bulkData, 2)
            If d(bulkData(i, 1))(j - 2) <> bulkData(i, j) Then
                tr.Failed = True
                tr.Message = "Dictionary data failed validation."
            End If
        Next j
    Next i

Finally:
    Set TestDictionary_AddBulkColMode = tr
End Function

Private Function TestDictionary_CountItems() As TestResult
Attribute TestDictionary_CountItems.VB_Description = "Tests the Count property of the dictionary."
'   Tests the Count property of the dictionary.
'
    Dim tr As New TestResult

'   Arrange
    Const ADDKEYS As String = "ABCDEFGHIJKLMNOP"

'   Act and Assert
    Dim d As New Dictionary
    If d.Count <> 0 Then
        tr.Failed = True
        tr.Message = "Expected 0 count but got " & d.Count
        GoTo Finally
    End If


    Dim i As Long
    For i = 1 To Len(ADDKEYS)
        d.Add Mid(ADDKEYS, i, 1), Nothing
        If d.Count <> i Then
            tr.Failed = True
            tr.Message = "Expected " & i & " count but got " & d.Count
            GoTo Finally
        End If
    Next i

Finally:
    Set TestDictionary_CountItems = tr
End Function

Private Function TestDictionary_ItemReturnsItem() As TestResult
Attribute TestDictionary_CountItems.VB_Description = "Tests the Count property of the dictionary."
'   Tests the Count property of the dictionary.
'
    Dim tr As New TestResult

'   Arrange
    Const EXPRESA As String = "A Result"
    Const EXPRESB As String = "B Result"
    Const INPKEYA As String = "A"
    Const INPKEYB As String = "B"

'   Act
    Dim d As New Dictionary
    d.Add INPKEYA, EXPRESA
    d.Add INPKEYB, EXPRESB

'   Assert
'   Test the Item method.
    Dim result As String
    result = d.Item(INPKEYA)
    If Not result = EXPRESA Then
        tr.Failed = True
        tr.Message = "Expected " & EXPRESA & " but got " & result
        GoTo Finally
    End If

'   Test the Item method (as default method).
    result = d(INPKEYB)
    If Not result = EXPRESB Then
        tr.Failed = True
        tr.Message = "Expected " & EXPRESB & " but got " & result
    End If

Finally:
    Set TestDictionary_ItemReturnsItem = tr
End Function