Attribute VB_Name = "Example"
' TinySeleniumVBA
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' �Z�b�g�A�b�v
'
' 1. �c�[�����Q�Ɛݒ��`Microsoft Scripting Runtime`���I���ɂ���
'
' 2. WebDriver.cls, WebElement.cls JsonConverter.bas ���v���W�F�N�g�ɒǉ�
'
' 3. WebDriver���_�E�����[�h�i�u���E�U�̃��W���[�o�[�W�����Ɠ������́j
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' �g����
'    `WebDriver`�̃C���X�^���X���_�E�����[�h����WebDriver���g���Đ������܂��B
'    ���������͉���Example���Q�Ƃ��������B
' ==========================================================================

' ==========================================================================
' Setup
'
' 1. Set reference to `Microsoft Scripting Runtime`
'
' 2. Add WebDriver.cls, WebElement.cls and JsonConverter.bas to your VBA Project
'
' 3. Download WebDriver (driver and browser should be the same version)
'   - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
'   - Chrome: https://chromedriver.chromium.org/downloads
'
' Usase
'    Create instance of `WebDriver` with the path to the driver you download.
'    See also the example below.
' ==========================================================================


' ==========================================================================
' Example
' ==========================================================================
Option Explicit

Public Sub main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    'Driver.Edge "path\to\msedgedriver.exe"
    
    'Shell "C:\Program Files\chromedriver\chromedriver.exe", vbMinimizedNoFocus
    
    Driver.Chrome "C:\Program Files\chromedriver\chromedriver.exe"

    
    ' Open browser
    Driver.OpenBrowser
    
    ' Navigate to Google
    Driver.Navigate "https://www.google.co.jp/?q=selenium"

    ' Get search textbox
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")
    
    ' Get value from textbox
    Debug.Print searchInput.GetValue
    
    ' Set value to textbox
    searchInput.SetValue "�����G���N�g����"
    
    Dim i, n As Long
    
    For i = 0 To 1000000
        n = n + 1
    Next
    
    
    
    Dim bttn
    Set bttn = Driver.FindElement(By.Name, "btnK")
    
    ' Click search button
    Driver.FindElement(By.Name, "btnK").Click
    
    ' Refresh - you can use Execute with driver command even if the method is not provided
    'Driver.Execute Driver.CMD_REFRESH
End Sub





