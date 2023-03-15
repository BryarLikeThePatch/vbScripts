'Okay this is fascinating and very powerful, so I decided to do a quick writup. I've discovered the ability to do JS injecitons in Visual Basic Selenium Driver. 
'For Example:

Sub jsInject()
  Dim table As Selenium.WebElement, tableLen As Selenium.WebElement
  Dim cDriver As Selenium.ChromeDriver
  Dim url As Stirng
  Dim t As Double
  Dim num As Integer, i As Integer
  
  t = Timer 'For benchmarking
  num = 2 'For the pasted row in Excel
  Application.ScreenUpdating 'For Speed
  
  url="https://TestURL.net"
  
  Set cDriver = New Selenium.ChromeDriver
    cDriver.Start
    cDriver.get(url)
      Set tableLen = cDriver.FindElementByName("results2_length")
        tableLen.AsSelect.SelectByValue("100")
        For i = 1 to 31
        Set table = cDriver.FindElementById("results2")
          table.AsTable.ToExcel ThisWorkbook.Worksheets("Sheet1").Range("A" & rNum)
          cDriver.ExecuteScript("docuement.querySelector('#results2_next').click();") '<= This is the crucial line. By using JS, I'm able to access functionality that SB was unable to find or execute. 
          num = num + 101 'for offseting
        Next i
End Sub
