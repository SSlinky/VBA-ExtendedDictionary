# Add method
Adds a key and item pair to the Scripting.Dictionary

## Syntax
_object_.**Add** _Key, _Val_

The **Add** method has the following parts:

Part | Description
--------- | ----------
_object_ | Required. Always the name of a **cExtendedDictionary** object.
_Key_ | Required. They key associated with the item being added.
_Val_ | Required. The value associated with the key being added.

## Remarks
If `OptionNoItemFail` is `False` then an error will be raised if the key already exists.

If the option is set to `True` then adding a key will not fail, it will override the previous value. This can be useful to avoid boilerplate checking for duplicates.