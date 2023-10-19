# OptionNoItemFail Property

Returns or sets a Boolean that configures the object to raise or ignore errors when getting or setting dictionary items.

## Syntax

_object_.**OptionNoItemFail**

Value | Action | Behaviour
------|--------|----------
False | Add item that exists | Exception is raised.
False | Get item that doesn't exist | Exception is raised.
True | Add item that exists | The value for that key is updated.
True | Get item that doesn't exist | Nothing is returned and no exception is raised.

## Remarks

This option allows safe getting and setting without having to write boilerplate to test whether the key exists or not.
