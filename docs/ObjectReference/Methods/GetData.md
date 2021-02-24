# Getdata method
Generates a Variant array with the key / value pairs in the Scripting.Dictionary.

## Syntax
_object_.**Getdata** _Key_, _[OptionUseRowMode]_

The **Getdata** method has the following parts:

Part                | Description
:---                | :---
_object_            | Required. Always the name of a **cExtendedDictionary** object.
_OptionUseRowMode_  | Optional. Orients keys and values vertically or horizontally.

## Remarks
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