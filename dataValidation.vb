Sub validate()
  Dim vRng As Range, r As Range
  Set vRng = ActiveSheet.Range("I2:I34")
  vRng.Validation.Add Type:=xlValidationList, AlertStyle:=xlValidAlertStop, Formula1:= "Fail,Pass" 'creates Validation with list values being either Fail or Pass
  For Each r in vRng
    If HAS_VALIDATION(r) = True Then 'checks if Validation is applied with UDF
      With r
        .Value = "Fail" 'sets default value as Fail
      End With 
    End If
End Sub

Function HAS_VALiDATION(c As Range) As Boolean
  Dim t: t = Null
  On Error Resume Next
  t = c.Validation.Type
  On Error GoTo 0
  HAS_VALIDATION = NOT IsNull(t)
End Function 
