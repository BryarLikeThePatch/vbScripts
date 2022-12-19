'There are times that you are unable to create automation to check the same conditions of Conditional Formatting and you're forced to check the 
'displayed color of the cell. The issue is that the Range.DisplayFormat.Interior.Color will not function will not work in UDF (user defined functions)
'If you try to call the DisplayFormat method, it will return a #Value error. To get around this, I've passed a string into a function that
'relies on a helper function to call the .DisplayFormat method and uses evaluate to get access to that value. 

Function CFColor(ByVal R As Range) As Double
    Application.Volatile (False)
    CFColor = Evaluate("Helper(" & R.Address() & ")")
End Function

Private Function Helper(ByVal R As Range) As Double 
    Helper = R.DisplayFormat.Interior.Color 
End Function