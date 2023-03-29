'The challenge was posed on Reddit's r/Excel community for the need of dividing a value equally between a specified number of rows. 
'The kicker was that the values had to be expressed to the nearest cent. 
'For example: If the value of $10.40 was to be divided between 5 individuals: it would be easy to return an array with $2.08 in each row. 
'However, in more complex examples, such as $10.13 between 5 individuals, you would end up with $2.026 in each row. 
'Instead, the user wanted the returned values to be distrubuted as fairly as possible without the need to divide cents. 
'So in the latter example: the user wanted the returned values to be an array of [2.03, 2.03, 2.03, 2.02, 2.02]
'This is the original solution that I came up with:


Function EQ(val As Double, num As INteger) As Variant()
    Dim n As Double, arr() As Variant, r As Double, digit As Integer
    digit = Len(CStr(val)) - InStr(CStr(val), ".") 'Find the number of digits after the decimal to help with rounding in the Private function
    n = val / num
    For i = 1 To num
        ReDim Preserve arr(1 To i)
            arr(i) = Round(n, 2)
    Next i
    i = 1
    If arrSum(arr, digit) > val Then
        Do Until arrSum(arr, digit) <= val
            arr(i) = arr(i) - 0.01
            i = i + 1
        Loop
    ElseIf arrSum(arr, digit) < val Then
        Do Until arrSum(arr, digit) >= val
            arr(i) = arr(i) + 0.01
            i = i + 1
        Loop
    End If
    EQUALITY = Application.WorksheetFunction.Transpose(arr)
End Function

Private Function arrSum(arr() As Variant, d As Integer) As Double
    Dim sum As Double
    For Each a In arr
        sum = sum + a
    Next a
    arrSum = Round(sum, d)
End Function

'To walk throguh what is happening:
' values passed in => ${val} -> the number to be divided "equally" and ${num} -> the number of indiivuals to be used as the denominator
' the ${val} is divided by ${num} and rounded to the nearest cent. 
' an array of length ${num} is constructed
' the array is then sent to a private function that acts like JavaScript's "array.reduce(() =>)" to produce a sum
' the sum is then sent back to the original function to a 'Do' loop. (VB's "While" loop equivilant)
' If the sum is greater than the original ${val}, it one cent is subtracted from ${i} number of variables until the array is ready to be returned
' On the flip side, if the sum is less than the original ${val}, one cent is incremented to ${i} number of variables in the array
'It then returns an array transposed to be vertical instead of horizontally displayed in Excel. 

'This can also be refactored to accept ranges (cell references) and extrapolating the values from the references

Function EQ_RFCTR(v As Range, x As Range) As Variant()
    Dim n As Double, arr() As Variant, r As Double, digit As Integer, val As Double, num As Integer
    val = v.Value
    num = x.Value
    digit = Len(CStr(val)) - InStr(CStr(val), ".")
    n = val / num
    For i = 1 To num
        ReDim Preserve arr(1 To i)
            arr(i) = Round(n, 2)
    Next i
    i = 1
    If arrSum(arr, digit) > val Then
        Do Until arrSum(arr, digit) <= val
            arr(i) = arr(i) - 0.01
            i = i + 1
        Loop
    ElseIf arrSum(arr, digit) < val Then
        Do Until arrSum(arr, digit) >= val
            arr(i) = arr(i) + 0.01
            i = i + 1
        Loop
    End If
    EQUALITY = Application.WorksheetFunction.Transpose(arr)
End Function

Private Function arrSum(arr() As Variant, d As Integer) As Double
    Dim sum As Double
    For Each a In arr
        sum = sum + a
    Next a
    arrSum = Round(sum, d)
End Function


'Finally for speed, we can refactor the original to keep the number of times the value was sent to the private function to 1. 
'The private function can also be replaced by the internal Sum function that can already accept arrays. 
'By assigning the original aggregate to a variable ${agg}, we can add +/-.01 to acheive the same result as before. 

Function EQ_RFCTR2(val As Double, num As INteger) As Variant()
  Dim n As Double, arr() As Variant, r As Double, digit As Integer, agg As Double
    digit = Len(CStr(val)) - InStr(CStr(val), ".") 'Find the number of digits after the decimal to help with rounding in the Private function
    n = val / num
    For i = 1 To num
        ReDim Preserve arr(1 To i)
            arr(i) = Round(n, 2)
    Next i
    i = 1
    agg = Round(Applicaiton.WorksheetFunction.Sum(arr),digit)
    If agg > val Then
        Do Until agg <= val
            arr(i) = arr(i) - 0.01
            i = i + 1
            agg = agg - 0.01
        Loop
    ElseIf agg < val Then
        Do Until agg >= val
            arr(i) = arr(i) + 0.01
            i = i + 1
            agg = agg + 0.01
        Loop
    End If
    EQUALITY = Application.WorksheetFunction.Transpose(arr)
End Function