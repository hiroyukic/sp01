Attribute VB_Name = "Module5_js"


'Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Sub Main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    Driver.Chrome "C:\Program Files\chromedriver\chromedriver.exe"



' Open browser
Driver.OpenBrowser

' Navigate to Google
'Driver.Navigate "https://www.google.co.jp/?q=liella"
Driver.Navigate "C:\tmp\sample.html"

' Show alert
'Driver.ExecuteScript "alert('Hello TinySeleniumVBA')"

' === Use breakpoint to CLOSE ALERT before continue ===

Sleep (1000)

' Pass argument
'Driver.ExecuteScript "alert('Hello ' + arguments[0] + ' as argument')", Array("TinySeleniumVBA")

' === Use breakpoint to CLOSE ALERT before continue ===

' Pass element as argument
Dim searchInput
'Set searchInput = Driver.FindElement(By.Name, "q")
'Driver.ExecuteScript "alert('Hello ' + arguments[0].value + ' ' + arguments[1])", Array(searchInput, "TinySeleniumVBA")

' === CLOSE ALERT and continue ===

' Get return value from script
Dim retStr As String
'retStr = Driver.ExecuteScript("return 'Value from script'")
Debug.Print retStr

' Get WebElement as return value from script
Dim firstDiv As WebElement
'Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('div')[0]")   '--OK
'Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('th')[0]")
'Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('div')[0]")

'Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('table')[0]")   '--ok

Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('tr')[0]")         '社名 ウェブサンプル株式会社

Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('th')[0]")         '社名
Debug.Print firstDiv.GetText()

Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('th')[1]")         '住所
Debug.Print firstDiv.GetText()

Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('td')[0]")         'ウェブサンプル株式会社
Debug.Print firstDiv.GetText()

Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('td')[1]")         '〒000-0000 東京都○○区○○○○○○
Debug.Print firstDiv.GetText()


' Get complex structure as return value from script
Dim retArray
retArray = Driver.ExecuteScript("return [['a', '1'], {'key1': 'val1', 'key2': document.getElementsByTagName('div'), 'key3': 'val3'}]")

Debug.Print retArray(0)(0)  ' a
Debug.Print retArray(0)(1)  ' 1

Debug.Print retArray(1)("key1") ' val1
Debug.Print retArray(1)("key2")(0).GetText()    ' Inner Text
Debug.Print retArray(1)("key2")(1).GetText()    ' Inner Text
Debug.Print retArray(1)("key3") ' val3


End Sub




