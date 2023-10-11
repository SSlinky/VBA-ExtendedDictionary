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

Sub RunSingle()
    Dim tr As TestResult
    Set tr = TestDictionary_ItemReturnsItem()
    
    If tr.Failed Then
        Debug.Print "FAIL: " & tr.Name, tr.Message
    Else
        Debug.Print "PASS"
    End If
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
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(1, d.Count) Then GoTo Finally
    If tr.AssertAreEqual(ADDVAL, d(ADDKEY)) Then GoTo Finally

Finally:
    On Error GoTo 0
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
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(Len(HDRS), d.Count) Then GoTo Finally

    For i = 1 To UBound(bulkData, 1)
        For j = 2 To UBound(bulkData, 2)
            If d(bulkData(i, 1))(j - 2) <> bulkData(i, j) Then
                tr.Failed = True
                tr.Message = "Dictionary data failed validation."
            End If
        Next j
    Next i

Finally:
    On Error GoTo 0
    Set TestDictionary_AddBulkColMode = tr
End Function

Private Function TestDictionary_AddBulkRowMode() As TestResult
Attribute TestDictionary_AddBulkRowMode.VB_Description = "Add bulk items to the dictionary."
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
    d.AddBulk Application.Transpose(bulkData), OptionUseRowMode:=True

'   Assert
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(Len(HDRS), d.Count, "count headers") Then GoTo Finally

    For i = 1 To UBound(bulkData, 1)
        For j = 2 To UBound(bulkData, 2)
            If d(bulkData(i, 1))(j - 2) <> bulkData(i, j) Then
                tr.Failed = True
                tr.Message = "Dictionary data failed validation."
            End If
        Next j
    Next i

Finally:
    On Error GoTo 0
    Set TestDictionary_AddBulkRowMode = tr
End Function

Private Function TestDictionary_AddBulkCountKeys() As TestResult
Attribute TestDictionary_AddBulkCountKeys.VB_Description = "Add bulk items to the dictionary."
'   Add bulk items to the dictionary.
    Dim tr As New TestResult

'   Arrange
    Dim bulkData() As Variant
    ReDim bulkData(1 To 6, 1 To 1)

    Const HDRS As String = "ABC"
    Const VALS As String = "123"

    Dim i As Long
    For i = 1 To Len(HDRS)
        Dim j As Long
        For j = 1 To CLng(Mid(VALS, i, 1))
            Dim n As Long
            n = n + 1
            bulkData(n, 1) = Mid(HDRS, i, 1)
        Next j
    Next i

    Dim d As New Dictionary

'   Act
    On Error Resume Next
    d.AddBulk bulkData, OptionCountKeys:=True

'   Assert
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(Len(HDRS), d.Count, "count keys") Then GoTo Finally

    For i = 1 To Len(HDRS)
        Dim h As String, v As Long
        h = Mid(HDRS, i, 1)
        v = CLng(Mid(VALS, i, 1))
        If tr.AssertAreEqual(v, d(h), h) Then Exit For
    Next i

Finally:
    Set TestDictionary_AddBulkCountKeys = tr
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
    If tr.AssertAreEqual(0, d.Count, "count keys") Then GoTo Finally

    Dim i As Long
    For i = 1 To Len(ADDKEYS)
        d.Add Mid(ADDKEYS, i, 1), Nothing
        If tr.AssertAreEqual(i, d.Count, "count keys at " & i) Then Exit For
    Next i

Finally:
    Set TestDictionary_CountItems = tr
End Function

Private Function TestDictionary_ItemReturnsItem() As TestResult
Attribute TestDictionary_ItemReturnsItem.VB_Description = "Tests the default and explicit item return."
'   Tests the default and explicit item return.
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
    If tr.AssertAreEqual(EXPRESA, d.Item(INPKEYA), INPKEYA) Then GoTo Finally
    If tr.AssertAreEqual(EXPRESB, d(INPKEYB), INPKEYB) Then GoTo Finally

Finally:
    Set TestDictionary_ItemReturnsItem = tr
End Function

Private Function TestDictionary_Exists() As TestResult
Attribute TestDictionary_Exists.VB_Description = "Tests Exists property works positively and negatively."
'   Tests Exists property works positively and negatively.
    Dim tr As New TestResult

'   Arrange
    Const INPKEYA As String = "A"
    Const INPKEYB As String = "B"
    Dim d As New Dictionary
    d.Add INPKEYA, Nothing

'   Act
    Dim posResult As Boolean
    posResult = d.Exists(INPKEYA)

    Dim negResult As Boolean
    negResult = d.Exists(INPKEYB)

'   Assert
    If tr.AssertIsTrue(posResult, "positive check") Then GoTo Finally
    If tr.AssertIsFalse(negResult, "negative check") Then GoTo Finally

Finally:
    Set TestDictionary_Exists = tr
End Function

Private Function TestDictionary_GetItemsReturnsAllItems() As TestResult
Attribute TestDictionary_GetItemsReturnsAllItems.VB_Description = "Test Items returns all items."
'   Test Items returns all items.
    Dim tr As New TestResult

'   Arrange
    Const EXPRESA As String = "A Result"
    Const EXPRESB As String = "B Result"
    Const INPKEYA As String = "A"
    Const INPKEYB As String = "B"
    
    Dim d As New Dictionary
    d.Add INPKEYA, EXPRESA
    d.Add INPKEYB, EXPRESB

'   Act
    Dim result As Variant
    result = d.Items()

'   Assert
    On Error Resume Next
    If tr.AssertAreEqual(EXPRESA, result(0), INPKEYA) Then GoTo Finally
    If tr.AssertAreEqual(EXPRESB, result(1), INPKEYB) Then GoTo Finally
    If tr.AssertNoException() Then GoTo Finally

Finally:
    Set TestDictionary_GetItemsReturnsAllItems = tr
End Function

Private Function TestDictionary_GetKeysReturnsKeys() As TestResult
Attribute TestDictionary_GetKeysReturnsKeys.VB_Description = "Test Keys returns all keys."
'   Test Keys returns all keys.
    Dim tr As New TestResult

'   Arrange
    Const INPKEYA As String = "A"
    Const INPKEYB As String = "B"
    
    Dim d As New Dictionary
    d.Add INPKEYA, Nothing
    d.Add INPKEYB, Nothing

'   Act
    Dim result As Variant
    result = d.Keys()

'   Assert
    On Error Resume Next
    If tr.AssertAreEqual(INPKEYA, result(0)) Then GoTo Finally
    If tr.AssertAreEqual(INPKEYB, result(1)) Then GoTo Finally
    If tr.AssertNoException() Then GoTo Finally

Finally:
    Set TestDictionary_GetKeysReturnsKeys = tr
End Function

Private Function TestDictionary_GetDataReturnsData() As TestResult
Attribute TestDictionary_GetDataReturnsData.VB_Description = "Test data out matches data in."
'   Test data out matches data in.
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

    On Error Resume Next
    Dim d As New Dictionary
    d.AddBulk bulkData

'   Act
    Dim results As Variant
    results = d.GetData()

'   Assert
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    For i = 1 To UBound(bulkData, 1)
        For j = 2 To UBound(bulkData, 2)
            If results(i, j) <> bulkData(i, j) Then
                tr.Failed = True
                tr.Message = "Dictionary data failed validation."
            End If
        Next j
    Next i

Finally:
    Set TestDictionary_GetDataReturnsData = tr
End Function

Private Function TestDictionary_OptionNoItemFailOverwrites() As TestResult
Attribute TestDictionary_OptionNoItemFailOverwrites.VB_Description = "OptionNoItemFail overwrites rather than throwing."
'   OptionNoItemFail overwrites rather than throwing.
    Dim tr As New TestResult

'   Arrange
    Const INPKEYA As String = "A"
    Const INPVALA As String = "A Value"
    Const INPVALB As String = "Visibly different value to A"
    
'   Act
    Dim d As New Dictionary
    d.OptionNoItemFail = True

    On Error Resume Next
    d.Add INPKEYA, INPVALA
    d.Add INPKEYA, INPVALB

'   Assert
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(1, d.Count, "count") Then GoTo Finally
    If tr.AssertAreEqual(INPVALB, d(INPKEYA), INPKEYA) Then GoTo Finally

Finally:
    Set TestDictionary_OptionNoItemFailOverwrites = tr
End Function

' TODO: Investigate!
'   Error handling in the Dictionary class seems to be messing with
'   error handling in the test scope. Exceptions thrown by the dictionary
'   do not seem to be caught (as they should) at this level.

' Private Function TestDictionary_NoOptionNoItemFailThrows() As TestResult
' Attribute TestDictionary_NoOptionNoItemFailThrows.VB_Description = "Without OptionNoItemFail throws rather than overwriting."
' '   Without OptionNoItemFail throws rather than overwriting.
'     Dim tr As New TestResult

' '   Arrange
'     Const DUPLICATEKEYEX As Long = 457
'     Const INPKEYA As String = "A"
'     Const INPVALA As String = "A Value"
'     Const INPVALB As String = "A Value"
    
' '   Act
'     Dim d As New Dictionary
'     d.OptionNoItemFail = False

'     d.Add INPKEYA, INPVALA
'     On Error Resume Next
'     d.Add INPKEYA, INPVALB

' '   Assert
'     If tr.AssertRaised(DUPLICATEKEYEX) Then GoTo Finally
'     On Error GoTo 0

'     If tr.AssertAreEqual(1, d.Count, "count") Then GoTo Finally
'     If tr.AssertAreEqual(INPVALA, d(INPKEYA), INPKEYA) Then GoTo Finally

' Finally:
'     Set TestDictionary_NoOptionNoItemFailThrows = tr
' End Function

Private Function TestDictionary_DataRowsAndColsCorrect() As TestResult
Attribute TestDictionary_DataRowsAndColsCorrect.VB_Description = "Tests the DataRows and DataCols properties."
'   Tests the DataRows and DataCols properties.
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

    On Error Resume Next
    Dim d As New Dictionary
    d.AddBulk bulkData

'   Act
    Dim dRowsColMode As Long
    dRowsColMode = d.DataRows()

    Dim dColsColMode As Long
    dColsColMode = d.DataCols()

    Dim dRowsRowMode As Long
    dRowsRowMode = d.DataRows(OptionUseRowMode:=True)

    Dim dColsRowMode As Long
    dColsRowMode = d.DataCols(OptionUseRowMode:=True)

'   Assert
    If tr.AssertNoException() Then GoTo Finally
    On Error GoTo 0

    If tr.AssertAreEqual(UBound(bulkData, 1), dRowsColMode) Then GoTo Finally
    If tr.AssertAreEqual(UBound(bulkData, 2), dColsColMode) Then GoTo Finally
    If tr.AssertAreEqual(UBound(bulkData, 2), dRowsRowMode) Then GoTo Finally
    If tr.AssertAreEqual(UBound(bulkData, 1), dColsRowMode) Then GoTo Finally

Finally:
    Set TestDictionary_DataRowsAndColsCorrect = tr
End Function