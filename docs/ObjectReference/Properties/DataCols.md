# DataCols Property

Returns the count of columns `GetData` would return.

## Syntax

_object_.**DataCols** (_[OptionUseRowMode]_)

Part                | Description
:---                | :---
_object_            | Required. Always the name of a **Dictionary** object.
_OptionUseRowMode_  | Optional. Whether to run in row mode or not.

## Remarks

This property is most useful when using `GetData` to write to a range.

The mode determines the shape of your data. If in column mode, with keys running down a column, this property returns the length of your data. Otherwise it returns the count of the keys.

`ed.DataCols` is equivalent to `ed.DataRows OptionUseRowMode:=True`
