# AddBulk method

Adds key value pairs from a 2D array. Supports keys as first row or firt column. Automatically detects array size and adds values based on number of values per key.

## Syntax

_object_.**AddBulk** _ValueArray2D_, _[OptionUseRowMode]_, _[OptionCountKeys]_

The **AddBulk** method has the following parts:

Part                    | Description
:---                    | :---
_object_                | Required. Always the name of a **Dictionary** object.
_ValueArray2D_          | Required. A two dimensional array of at least one row and column.
_[OptionUseRowMode]_    | Optional. Use the first row instead of the first column as keys.
_[OptionCountKeys]_     | Optional. The value is the number of times the key has been found. This will force `OptionNoItemFail` to True.

## Examples

### Key only

Load a simple 2D array with only one column. The values will be defaulted to Nothing.

```vba
Dim ed As New Dictionary
ed.AddBulk Range("A1:A50").Value
```

### Key / single value pairs

Load a simple 2D array where the first column is the key and the second column is the value.

```vba
Dim ed As New Dictionary
ed.AddBulk Range("A1:B50").Value
```

### Horizontal key / multiple value pairs

Load a 2D array with a key and more than one value. Using `OptionUseRowMode` we can specify that the keys are in the first row of the
array rather than the default first column behaviour.

```vba
Dim ed As New Dictionary
ed.AddBulk Range("A1:Z5").Value, OptionUseRowMode=True
```

### Count unique values

Load simple 2D array. Only the first row or column will be considered, depending on `OptionUseRowMode`. The values will be the count of times the key
appears in the passed in array.

```vba
Dim ed As New Dictionary
ed.AddBulk Range("A1:A500").Value, OptionCountKeys=True
```

The use of `OptionCountKeys` implies [OptionNoItemFail](./ObjectReference/Properties/OptionNoItemFail.md). Using this option will override the property to `True`.
