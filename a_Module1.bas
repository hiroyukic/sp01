Attribute VB_Name = "Module1"
Public Sub main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    Driver.Edge "path\to\msedgedriver.exe"
    
    C:\Program Files\chromedriver

    ' Open browser
    Driver.OpenBrowser

    ' Navigate to Google
    Driver.Navigate "https://www.google.co.jp/?q=cat"

    ' Get search textbox
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")

    ' Get value from textbox
    Debug.Print searchInput.GetValue

    ' Set value to textbox
    searchInput.SetValue "”L ƒTƒoƒgƒ‰”’"

    ' Click search button
    Driver.FindElements(By.Name, "btnK")(1).Click
End Sub
