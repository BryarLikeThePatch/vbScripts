'ArrayToString

Function arrayToString(arr As Variant()) As String
    Dim s As String
    For Each v in arr
        s = s & "," & v
    Next v
End Function 

