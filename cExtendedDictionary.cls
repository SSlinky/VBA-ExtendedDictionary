VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ------------------------------------------------
'   Dictionary wrapper class to extend standard
'   functionality and improve error handling
'
'   Author : Sam Vanderslink
'   Web    : notis.net.au
'   Licence: GNU GPLv3 (open source)
' ------------------------------------------------


' ------------------------------------------------
'   Enums - must be at top of module
' ------------------------------------------------

'A value representing the comparison mode used by functions such
'as StrComp. Can only be changed while the dictionary has no data.
Enum CompareModeType
'   Performs a comparison by using the setting of the Option Compare statement.
    vbUseCompareOption = -1
'   Performs a binary comparison.
    vbBinaryCompare = 0
'   Performs a textual comparison.
    vbTextCompare = 1
'   Microsoft Access only. Performs a comparison based on information
'   in your database. This option is included for completeness only.
    vbDatabaseCompare = 2
End Enum


' ------------------------------------------------
'   Private Properties
' ------------------------------------------------

' The underlying dictionary object
Private mBaseDict As Object

' Whether an duplicative item add should raise an error
' or not. Set True to prevent duplicates raising an error.
Private mOptionNoItemFail As Boolean

' The largest value array.
Private mDataShapeValues


' ------------------------------------------------
'   Constructor
' ------------------------------------------------

Private Sub Class_Initialize()
    Set mBaseDict = CreateObject("Scripting.Dictionary")
End Sub


' ------------------------------------------------
'   Public Properties
' ------------------------------------------------

'Prevents .Add from failing when the key already exists.
'Prevents .Item from failing when the key doesn't exist.
Public Property Get OptionNoItemFail() As Boolean
    OptionNoItemFail = mOptionNoItemFail
End Property

Public Property Let OptionNoItemFail(val As Boolean)
    mOptionNoItemFail = val
End Property

'Sets the CompareMode of the underlying Scripting.Dictionary.
'See CompareModeType for more information.
Public Property Get CompareMode() As CompareModeType
    CompareMode = mBaseDict.CompareMode
End Property

Public Property Let CompareMode(CompareM As CompareModeType)
    mBaseDict.CompareMode = CompareM
End Property

'Gets the key count of the underlying Scripting.Dictionary.
Public Property Get Count() As Variant
    Count = mBaseDict.Count
End Property

'Sets or returns an item for a specified key.
Public Property Get Item(Key As Variant) As Variant
    If mOptionNoItemFail Then
        If mBaseDict.Exists(Key) Then
            Item = mBaseDict.Item(Key)
        End If
    Else
        Item = mBaseDict.Item(Key)
    End If
End Property

Public Property Let Item(Key As Variant, val As Variant)
    If mOptionNoItemFail Then On Error Resume Next
    mBaseDict.Item(Key) = val
    If Err = 0 Then SetMetaData val
    On Error GoTo 0
End Property

'Updates a key, replacing the old key with the new key.
Public Property Let Key(Key As Variant, NewKey As Variant)
    mBaseDict.Key(Key) = NewKey
End Property

'Wrapper property that returns the value for the specified key if it exists.
'If it doesn't exist, it returns the default rather than raise an error.
Public Property Get GetValue(Key As Variant, ItemDefault As Variant) As Variant
    If mBaseDict.Exists(Key) Then
        GetValue = mBaseDict.Item(Key)
    Else
        GetValue = ItemDefault
    End If
End Property


'Returns the row count that would be returned by the GetData method
'The values are always +1 to account for the first row being keys
Public Property Get DataRows(Optional OptionUseRowMode As Boolean) As Long
'   OptionUseRowMode=True:  Returns the longest array size of values
'                           plus row for keys (if they exist)
'   OptionUseRowMode=False: Returns the count of the keys
    DataRows = IIf(OptionUseRowMode, mDataShapeValues + Min(1, Me.Count), Me.Count)
End Property

'Returns the column count that would be returned by the GetData method
'The values are always +1 to account for the first col being keys
Public Property Get DataCols(Optional OptionUseRowMode As Boolean) As Long
'   OptionUseRowMode=True:  Returns the count of the keys
'   OptionUseRowMode=False: Returns the longest array size of values
'                           plus 1 col for keys (if they exist)
    DataCols = IIf(OptionUseRowMode, Me.Count, mDataShapeValues + Min(1, Me.Count))
End Property


' ------------------------------------------------
'   Public Methods
' ------------------------------------------------

'Adds a key and item pair to the Scripting.Dictionary
Public Sub Add(Key As Variant, val As Variant)
    If mOptionNoItemFail Then On Error Resume Next
    mBaseDict.Add Key, val
    
    If Err = 457 Then
        Me.Item(Key) = val
    Else: SetMetaData val
    End If
    
    On Error GoTo 0
End Sub

'Removes a key/item pair from the Scripting.Dictionary.
Public Sub Remove(Key As Variant)
    Dim val As Variant
    val = Me.Item(Key)
    mBaseDict.Remove Key
    UntrackLargestValue val
End Sub

'The RemoveAll method removes all key, item pairs from the Scripting.Dictionary.
Public Sub RemoveAll()
    mBaseDict.RemoveAll
    mDataShapeValues = 0
End Sub

'Adds key value pairs from a 2D array. Supports keys as first row or firt column.
'Automatically detects array size and adds values based on number of values per key.
'   0 : Keys only (values are Nothing)
'   1 : Single values only
'  >1 : Values are a 1D array of values
Public Sub AddBulk(ValueArray2D As Variant, _
                    Optional OptionUseRowMode As Boolean, _
                    Optional OptionCountKeys As Boolean)
'       OptionUseRowMode : Use the first row instead of column as keys
'       OptionCountKeys  : The value is the number of times the key has been found
'                          This will force OptionNoItemFail to True

    Dim i As Long       ' For loop tracking
    Dim j As Long       ' For loop tracking
    Dim r As Long       ' Number of rows
    Dim c As Long       ' Number of columns
    Dim k As Variant    ' Dictionary key
    Dim v As Variant    ' Dictionary value
    
    If OptionCountKeys Then mOptionNoItemFail = True
    
    c = UBound(ValueArray2D, 2)
    r = UBound(ValueArray2D, 1)
    
    If OptionUseRowMode Then
'       Keys are column headers, add values by rows
        For i = 1 To c
            k = ValueArray2D(1, i)
            If OptionCountKeys Then
'               Add a count for this key
                v = Me.GetValue(k, 0) + 1
            Else
'               Add the values by row for header
                Select Case r
                    Case 2:
'                       Add singular value
                        v = ValueArray2D(2, i)
                    Case Is > 2:
'                       Add array of values from rows
                        ReDim v(r - 2)
                        For j = 0 To r - 2
                            v(j) = ValueArray2D(j + 2, i)
                        Next j
                    Case Else
'                       Default (no values)
                        v = Nothing
                End Select
            End If
            Me.Add k, v
        Next i
    Else
'       Keys are row headers, add values by columns
        For i = 1 To r
            k = ValueArray2D(i, 1)
            If OptionCountKeys Then
                v = Me.GetValue(k, 0) + 1
            Else
'               Add the values by column for header
                Select Case c
                    Case 2:
'                       Add singular value
                        v = ValueArray2D(i, 2)
                    Case Is > 2:
'                       Add array of values from columns
                        ReDim v(c - 2)
                        For j = 0 To c - 2
                            v(j) = ValueArray2D(i, j + 2)
                        Next j
                    Case Else
'                       Default (no values)
                        Set v = Nothing
                End Select
            End If
            Me.Add k, v
        Next i
    End If
End Sub


' ------------------------------------------------
'   Public Functions
' ------------------------------------------------

'Returns True if a specified key exists in the Scripting.Dictionary
'object; False if it does not.
Public Function Exists(Key As Variant) As Boolean
    Exists = mBaseDict.Exists(Key)
End Function

'Returns an array containing all the items in the Scripting.Dictionary.
Public Function Items() As Variant()
    Items = mBaseDict.Items
End Function

'Returns an array containing all existing keys in the Scripting.Dictionary.
Public Function Keys()
    Keys = mBaseDict.Keys
End Function

'Returns the dictionary data in a similar fashion to how it is
'loaded in with AddBulk.
Public Function GetData(Optional OptionUseRowMode As Boolean) As Variant
    Dim r As Long, c As Long
    
    Dim arrKeys     As Variant
    Dim arrVals     As Variant
    Dim results()   As Variant
    Dim keyRows     As Long
    Dim valCols     As Long
    
    arrKeys = mBaseDict.Keys()
    arrVals = mBaseDict.Items()
    keyRows = UBound(arrKeys) + 1
    
    keyRows = Me.DataRows()
    valCols = Me.DataCols()
    
'   Set up results array in the shape of the keys and values
'   adding an additional column to fit the keys as a column
    ReDim results(1 To keyRows, _
                  1 To valCols)

    
'   Populate the row(s) and column(s)
    For r = 1 To keyRows
        results(r, 1) = arrKeys(r - 1)
        If valCols > 2 Then
'           Populate columns for multi-column
            For c = 2 To valCols
                results(r, c) = arrVals(r - 1)(c - 2)
            Next c
        ElseIf valCols = 2 Then
'           Populate single value
            results(r, 2) = arrVals(r - 1)
        End If
    Next r
    
    If OptionUseRowMode Then
        GetData = Application.Transpose(results)
    Else
        GetData = results
    End If
    
End Function


' ------------------------------------------------
'   Private Helpers
' ------------------------------------------------


'Returns the number of dimensions in the passed in array
'This is a hack that intentionally seeks to raise an
'exception as there is no property or elegant way to do this.
Private Function NArrayDimensions(arr) As Long
    Dim i As Long
    
'   Test array dimensions until exception raised
    On Error GoTo ReportResult
    Do
        i = i + 1
        NArrayDimensions = UBound(arr, i)
    Loop
    
ReportResult:
    Err.Clear
    NArrayDimensions = i - 1
End Function

'Sets the metadata associated with the value
Private Sub SetMetaData(val As Variant)
    Dim valSize As Long
    
'   Get the size of the value or array
    valSize = GetValueOrArraySize(val)
    
'   Set the longest value if it is longer
    If valSize > mDataShapeValues Then _
        mDataShapeValues = valSize
End Sub

'Untracks the largest sized value array and resets
'to the next smallest size array.
Private Sub UntrackLargestValue(remVal As Variant)
    Dim val         As Variant
    Dim valSize     As Long
    Dim remValSize  As Long


'   Handle last value being removed
    If Me.Count = 0 Then
        mDataShapeValues = 0
        Exit Sub
    End If
    
'   Get the size of the value or array being removed
    remValSize = GetValueOrArraySize(remVal)
    
'   Reset the internal tracking
    mDataShapeValues = 0
    
'   Find the next largest item
    For Each val In Me.Items
        valSize = GetValueOrArraySize(val)
        
'       Next largest is the same size
        If valSize = remValSize Then
            mDataShapeValues = valSize
            Exit Sub
        End If
        
'       All other values
        If valSize > mDataShapeValues Then _
            mDataShapeValues = valSize
    Next val
End Sub

'Returns the length of the array or 1 for a value
Private Function GetValueOrArraySize(val) As Long
'   Handle the value being Nothing
    If TypeName(val) = "Nothing" Then Exit Function
    
'   Get the size of the array or 1 if not an array
    Select Case NArrayDimensions(val)
        Case Is = 0: GetValueOrArraySize = 1
        Case Is = 1: GetValueOrArraySize = UBound(val) + 1
    End Select
End Function

'Returns the smaller of the two passed in values
Private Function Min(a, b)
    Min = IIf(a < b, a, b)
End Function
