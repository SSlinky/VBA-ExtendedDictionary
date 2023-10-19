# DataRows Property

Returns the count of rows `GetData` would return.

## Syntax

_object_.**DataRows** (_[OptionUseRowMode]_)

Part                | Description
:---                | :---
_object_            | Required. Always the name of a **Dictionary** object.
_OptionUseRowMode_  | Optional. Whether to run in row mode or not.

## Remarks

This property is most useful when using `GetData` to write to a range.

The mode determines the shape of your data. If in column mode, with keys running down a column, this property returns the count of the keys. Otherwise it returns the length of your data.

`ed.DataRows` is equivalent to `ed.DataCols OptionUseRowMode:=True`
