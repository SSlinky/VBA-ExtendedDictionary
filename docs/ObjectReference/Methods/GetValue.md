# GetValue method

Wrapper property that returns the value for the specified key if it exists.
If it doesn't exist, it returns the default rather than raise an error.

## Syntax

_object_.**GetValue** _Key_, _ItemDefault_

The **GetValue** method has the following parts:

Part            | Description
:---            | :---
_object_        | Required. Always the name of a **Dictionary** object.
_Key_           | Required. They key associated with the value being looked up.
_ItemDefault_   | Required. The default value to be returned if the key is not found.

## Remarks

This method has been added to provide an equivalent getter to those found in other modern languages.

## Examples

### Key value counts

```vba
Dim cityCounter As New Dictionary
cityCounter.AddBulk Range("A1:A500").Value, OptionCountKeys=True

' Assume the key Perth appeared in the data 5 times but Melbourne
' didn't appear at all. Attempting to get the count for Melbourne
' would result in an error but the desired result is 0.

Debug.Print "Perth: " & cityCounter.GetValue("Perth", 0)        ' Perth: 17
Debug.Print "Perth: " & cityCounter.GetValue("Melbourne", 0)    ' Melbourne: 0
```
