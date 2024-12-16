# VBScript GetObject Function Bug

This repository demonstrates a common, yet subtle, error in VBScript's GetObject function when used with error handling. The bug arises from improper handling of the `On Error Resume Next` statement.

## Bug Description

The `GetObject` function attempts to retrieve an object from a collection.  It uses `On Error Resume Next` to suppress errors if the object isn't found. However, if an error does occur, the `obj` variable remains `Nothing`, leading to potential runtime errors later in the code.

## Solution

The solution involves explicitly checking if the object was successfully retrieved *after* the `On Error Resume Next` block.  This ensures that subsequent code operations don't act on an uninitialized variable.

## How to Reproduce

1. Run the `bug.vbs` script.
2. Observe the error message if the object doesn't exist.
3. Run the `bugSolution.vbs` script. Note the absence of an error when the object is not found. 