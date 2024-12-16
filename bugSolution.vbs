Function GetObject(objName)
'This improved function explicitly checks for errors after the On Error Resume Next block
Dim obj As Object
On Error Resume Next
Set obj = myCollection(objName)
On Error GoTo 0

If Err.Number <> 0 Then
    ' Handle the error appropriately
    Err.Clear 'Clear the error object
    Set GetObject = Nothing 'Return Nothing to indicate failure
    Exit Function
End If

Set GetObject = obj
End Function