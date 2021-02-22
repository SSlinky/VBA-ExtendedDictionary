VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "cExtendedDictionary"
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
'   Private Properties
' ------------------------------------------------

' The underlying dictionary object
Private mBaseDict As Object

' Whether an duplicative item add should raise an error
' or not. Set True to prevent duplicates raising an error.
Private mOptionNoItemFail As Boolean


' ------------------------------------------------
'   Constructor
' ------------------------------------------------

Private Sub Class_Initialize()
    Set mBaseDict = CreateObject("Scripting.Dictionary")
End Sub


' ------------------------------------------------
'   Public Properties
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

'Prevents .Add from failing when the key already exists.
'Prevents .Item from failing when the key doesn't exist.
Public Property Get OptionNoItemFail() As Boolean
    OptionNoItemFail = mOptionNoItemFail
End Property

Public Property Let OptionNoItemFail(Val As Boolean)
    mOptionNoItemFail = Val
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

Public Property Let Item(Key As Variant, Val As Variant)
    If mOptionNoItemFail Then On Error Resume Next
    mBaseDict.Item(Key) = Val
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


' ------------------------------------------------
'   Public Methods
' ------------------------------------------------

'Adds a key and item pair to the Scripting.Dictionary
Public Sub Add(Key As Variant, Val As Variant)
    If mOptionNoItemFail Then On Error Resume Next
    mBaseDict.Add Key, Val
    If Err = 457 Then Me.Item(Key) = Val
    On Error GoTo 0
End Sub

'Removes a key/item pair from the Scripting.Dictionary.
Public Sub Remove(Key As Variant)
    mBaseDict.Remove Key
End Sub

'The RemoveAll method removes all key, item pairs from the Scripting.Dictionary.
Public Sub RemoveAll()
    mBaseDict.RemoveAll
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
            mBaseDict.Add k, v
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
Public Function Exists(Key As String) As Boolean
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
