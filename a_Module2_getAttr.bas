Attribute VB_Name = "Module2_getAttr"
' ==========================================================================
'
' GetAttribute
'  2022/01/29
'
' ==========================================================================
Public Sub main()
  On Error GoTo error_label
  
  'Browser ‹N“®
    Set objDrv = New WebDriver
    'objDrv.Chrome "chromedriver.exe"
    objDrv.Chrome "C:\Program Files\chromedriver\chromedriver.exex"
    objDrv.OpenBrowser


    'URL‚ðŠJ‚­
    objDrv.Navigate Application.ThisWorkbook.Path & "\" & "GetAttr.html"

    'User
    Debug.Print "User GetAttribute(""value"")"
    Debug.Print vbTab & objDrv.FindElement(By.ID, "user01").GetAttribute("value")
    Debug.Print "User GetOuterHTML"
    Debug.Print vbTab & objDrv.FindElement(By.ID, "user01").GetOuterHTML

    'Passwoed
    Set objElm = objDrv.FindElement(By.ID, "pass01")
    Debug.Print "Passwoed GetAttribute(""value"")"
    Debug.Print vbTab & objElm.GetAttribute("value")
    Debug.Print "Passwoed GetOuterHTML"
    Debug.Print vbTab & objElm.GetOuterHTML

    'href
    Set objElm = objDrv.FindElement(By.ID, "YAHOO")
    Debug.Print "href GetAttribute(""href"")"
    Debug.Print vbTab & objElm.GetAttribute("href")

    'div
    Debug.Print "div GetInnerHTML"
    Debug.Print vbTab & objDrv.FindElement(By.TagName, "div").GetInnerHTML
    Debug.Print "div GetAttribute(""innerHTML"")"
    Debug.Print vbTab & objDrv.FindElement(By.TagName, "div").GetAttribute("innerHTML")

    'div
    Debug.Print "div GetOuterHTML"
    Debug.Print vbTab & objDrv.FindElement(By.TagName, "div").GetOuterHTML
    Debug.Print "div GetAttribute(""outerHTML"")"
    Debug.Print vbTab & objDrv.FindElement(By.TagName, "div").GetAttribute("outerHTML")

error_label:
    'Browser Close
    objDrv.CloseBrowser
    objDrv.Shutdown


End Sub

