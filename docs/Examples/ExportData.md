# Export Data to Range

Exporting the data is nearly as simple as importing it with `AddBulk`.

Using the `DataRows` and `DataCols` properties, we can resize a range so we can set the values to the output of `GetData`.

```vba
Sub ExportExample(ed As Dictionary, dst As Range)
    dst.Resize(ed.DataRows, ed.DataCols).Value = ed.GetData
End Sub
```
