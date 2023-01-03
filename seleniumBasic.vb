'If you want to do basic (no pun intended) web scraping using Excel for Desktop, SeleniumBasic is your best option without having to compile C# or know Python. 
'First, download the GitHub Repository for Selenium Basic here: https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0
'Next, update your chrome driver to match your current driver at the following: https://chromedriver.chromium.org/downloads
'Finally add Selenium to your references under the 'Tools' dropdown menu in VB. 
'This is a quick example of a very simple Selenium Code in VB. 

'Define SubRoutine and add boilerplate dimensions to worksheet
Option Explicit
Sub firstSeleniumCode()
  Dim rNum as Integer, lRow as Integer
  Dim ws As Worksheet
    Set ws = ActiveSheet 'Assign activesheet to ws object dimension
    lRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'find last row by counting from the bottom up to the first non-blank row
End Sub

'Reminder: for objects like Worksheets and Ranges, you have to use the 'Set' method to assign an object to the dimenstion. 
'For simple data types like strings, integers, and floats you can assign values directly without requiring the 'Set' method. 

'Next, we can outline our web objects, URLs, and chrome instance as dimensions. 

Option Explicit
Sub firstSeleniumCode()
  Dim rNum as Integer, lRow as Integer
  Dim ws As Worksheet
  'Selenium Dimensions
    Dim chromeInst As Selenium.ChromeDriver
    Dim url as String
    Dim dataTable as Selenium.WebElement
'assign variables
  Set ws = ActiveSheet 'Assign activesheet to ws object dimension
  lRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'find last row by counting from the bottom up to the first non-blank row
End Sub

'Now that the basics are added to the code, it's time to add the final pull using a queryselector

Option Explicit
Sub firstSeleniumCode()
  Dim rNum as Integer, lRow as Integer
  Dim ws As Worksheet
  Dim t as Timer
  'Selenium Dimensions
    Dim chromeInst As Selenium.ChromeDriver
    Dim url as String
    Dim dataTable as Selenium.WebElement
'assign variables
  Set ws = ActiveSheet 'Assign activesheet to ws object dimension
  lRow = ws.Range("A" & Rows.Count).End(xlUp).Row 'find last row by counting from the bottom up to the first non-blank row

'set URL
  url = "https://en.wikipedia.org/wiki/Ice_climbing"
'set chromedriver as new driver
  Set chromeInst = new ChromeDriver
    chromeInst.Start
'Navigate to the url
  chromeInst.Get(url)
  'add a short wait to ensure that the page has loaded
    t = Timer
    Do Until Timer - t > 10
    Loop
  'Find element with class and return to Excel. 
    rNum = 2
      set dataTable = chromeInst.FindElementByClass("wikitable floatright") 
        dataTable.AsTable.ToExcel(ws.Range("A" & rNum)
    
    msgBox "Web Page Has Ben Scraped"
    
End Sub
