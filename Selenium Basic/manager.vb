Option Explicit

Sub managers()
    Dim ws As Worksheet
    Dim names() As Variant, name As Variant, v As String
    Dim lRow As Integer, r As Integer
    'Selenium dimensions
        Dim cd As Selenium.ChromeDriver
        Dim url As String, suffix As String, selector As String
        Dim wrapper As Selenium.WebElements
        Dim manager As Selenium.WebElement
  'array of names
    Set ws = ActiveSheet
    lRow = ws.Range("G" & Rows.Count).End(xlUp).Row
    
    For r = 1 To lRow
      ReDim Preserve names(1 To r)
      name = ws.Range("G" & r).Value
      names(r) = name
    Next r
  
  'launch chromedriver and login
    Set cd = New ChromeDriver
      cd.Start
    url = "https://internalhierarchy.com"
    cd.Get (url)
    'Wait

  'Loop through name array and extract the manager from the Hierarchy
    url = "https://internalhierarchy.com/?u="
    suffix = "&v=work"
    
    r = 1
    For Each name In names
      On Error Resume Next
      cd.Get (url + name + suffix)
      Set wrapper = cd.FindElementByCss("[aria-label^=""Organization - this shows a hierarchy""]").FindElementsByCss("[class^=Persona-module__clickableUserContainer]")
      selector = wrapper(wrapper.Count - 1).Text
      With ws.Range("H" & r)
        .Value = selector
        .WrapText = False
      End With
      r = r + 1
      selector = ""
    Next name

      
End Sub
