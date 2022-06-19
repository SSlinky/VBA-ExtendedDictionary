# GetData method
Generates a Variant array with the key / value pairs in the Scripting.Dictionary.

## Syntax
_object_.**GetData** (_[OptionUseRowMode]_)

The **GetData** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. Always the name of a **Dictionary** object.
_OptionUseRowMode_  | Optional. Orients keys and values vertically or horizontally.

## Remarks
This function returns a base 1 2D array of type variant. This means it is already in an appropriate format to insert into a range.

The shape of the array will depend on the data. They keys will define the size of the first or second dimension, depending on whether `OptionUseRowMode` is True or not. The values will populate the other dimension if they exist.

Note that if no data exists in the dictionary, the returned array will be a single cell with Nothing.

## Examples
### Using OptionUserRowMode

`OptionUseRowMode:=True` keys are arranged across columns in the first row. Values populate rows below their respective key.

| k1  | k2  | k3  |
| --- | --- | --- |
| v1a | v2a | v3a |
| v1b | v2b | v3b |
| v1c | v2c | v3c |

`OptionUseRowMode:=False` keys are arranged across rows in the first column. Values populate columns right of their respective key.

|        |     |     |     |
| ---    | --- | --- | --- |
| **k1** | v1a | v1b | v1c |
| **k2** | v2a | v2b | v2c |
| **k3** | v3a | v3b | v3c |

### Returned Array Sizes
Array sizes returned by data type. Assumes `OptionUseRowMode` is `False`. Note that in these cases, the first column is the key so
the values start from the second column.

Keys    | Values per Key    | Return Shape (D1, D2)
---     | ---               | :---
0       | 0                 | (1 to 1, 1 to 1)
10      | 0                 | (1 to 10, 1 to 1)
16      | 255               | (1 to 16, 1 to 256)