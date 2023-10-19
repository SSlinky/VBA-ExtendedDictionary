# Counting value occurrences

Based on [this question](https://www.reddit.com/r/vba/comments/lp5vxz/vba_code_to_count_for_certain_instances_of_text/) asked on Reddit.

## Problem statement

Data exists in column A. Occurrences of the words "INAUDIBLE" and "NO RESPONSE" must be counted if the cell is not hidden.

## Solution

ExtendedDictionary (ed) makes this trivial.

1. Declare variablesa and define the data range.
2. Add the data to ed using the `AddBulk` method with `OptionCountKeys` enabled.
3. Use the `GetValue` method to safely retrieve a value, or return the deafult 0 if it doesn't exist.

```vba
Option Explicit

Sub CountInstances()

'   Variable declarations
    Dim arr As Variant
    Dim rg  As Range
    Dim k   As Variant
    Dim r   As Range
    
    Dim ed As New Dictionary
    
'   Arrange the visible data into an array
    Set rg = Range("A1:A" & Range("A" & Rows.Count).End(xlUp).Row)
    
'   Each contiguous batch of cells must be iterated through
'   since .SpecialCells returns a union and .Value gets the
'   values from the first range in the union only.
    For Each r In rg.SpecialCells(xlCellTypeVisible, True).Areas
'       Add the visible cells to the dictionary, counting the keys.
        ed.AddBulk r.Value, OptionCountKeys:=True
    Next r
    
'   Print the count values to the immediate window.
    For Each k In Array("INAUDIBLE", "NO RESPONSE")
        Debug.Print k & ":", ed.GetValue(k, 0)
    Next

End Sub
```
