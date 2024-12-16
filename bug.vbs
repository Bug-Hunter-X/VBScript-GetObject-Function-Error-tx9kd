Function GetObject() 
'This function is supposed to retrieve an object from a collection based on its name.
'However, it has a subtle bug that can lead to unexpected errors.

Dim obj As Object

On Error Resume Next
Set obj = myCollection(objName)
On Error GoTo 0

If obj Is Nothing Then
    Err.Raise vbObjectError + 1, , "Object not found"
End If

Set GetObject = obj
End Function