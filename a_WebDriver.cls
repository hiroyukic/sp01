VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "WebDriver"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False


' TinySeleniumVBA v0.1.0
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2021 uezo
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================
'
' ==========================================================================
' Copyright (c) 2021-2022 ishi
' The license complies with the "MIT License" above.
' --------------------------------------------------------------------------
' [Change History]
' 2021/06/17- v0.1.0a  ishi  Add : Frame, Alert, Select
' 2021/06/23  v0.1.0b  ishi  Add : Window
' 2021/06/24  v0.1.0c  ishi  Add : Timeout, Other
'                            Chg : Execute
' 2021/06/25  v0.1.0d  ishi  Add : Switch
' 2021/06/25  v0.1.0e  ishi  Add : GetAttribute
' 2021/06/28  v0.1.0f  ishi  Add : Close Window
' 2021/06/30  v0.1.0g  ishi  Add : innerHTML, outerHTML
' 2021/07/04  v0.1.0h  ishi  Add : IsElementPresent, ExecuteScript
'                            Add : Wait, Clear
' 2021/07/09  v0.1.0i  ishi  Add : GetCurrentFrame
' 2021/07/10  v0.1.0j  ishi  Add : SetFocus, SetFocusNoScroll
'                            Add : GetWindowSize, SetWindowSize
'                            Add : GetWindowPosision, SetWindowPosision
' 2021/07/13  v0.1.0k  ishi  Add : IsEnabled, IsDisplayed, GetScreenSize
' 2021/07/14  v0.1.0l  ishi  Add : GetImplicitlyWait, GetPageLoadTimeout
'                            Add : SetScriptTimeout, GetSetScriptTimeout
'                            Chg : IsElementPresent, Execute
'                            Add : MaximizeWindow, FullScreenWindow, MinimizeWindow
' 2021/07/15  v0.1.0m  ishi  Add : By.XPath, By.LinkText, By.PartialLinkText
' 2021/07/16  v0.1.0n  ishi  Add : GetCurrentWindowHandle
' 2021/07/20  v0.1.0o  ishi  Add : getAttribute_js (python 3.9.6 selenium 3.141.0)
'                            Chg : ExecuteScript
'                            Chg : GetAttribute, innerHTML, outerHTML
' 2021/07/24  v0.1.0p  ishi  Add : Scroll
'                            Chg : Edge, Chrome, Start (VbAppWinStyle Add)
' 2021/07/28  v0.1.0q  ishi  Chg : Driver commands Public -> Private
' 2021/08/07  v0.1.0r  ishi  Add : Back, Forward, Refresh,
'                                  GetActiveElement, CurrentURL, GetProperty
'                            Chg : Edge, Chrome, Start
' 2021/08/10  v0.1.0s  ishi  Chg : Edge, Chrome ( Enum AppWinStyle )
'                            Chg : SelectByValue
'                            Add : SelectByVisibleText, SelectByIndex
'                            Add : SelectTagsOptions
'                            Add : SelectedOptionText, AllSelectedOptions
'                            Add : DeSelectAll, DeSelectByValue
'                            Add : DeSelectByVisibleText, DeSelectByIndex
' 2021/08/13  v0.1.0t  ishi  Add : Enum Keys
'                            Chg : SetValue
' 2021/08/14  v0.1.0u  ishi  Add : SaveScreen
' 2021/08/22  v0.1.0v  ishi  Add : PageSource
' 2021/09/23  v0.1.0w  ishi  Add : Browser Options
'                                    CrateOptions, addArgument, addOption
'                                    BuildCapabilities, BuildDesiredCapabilities
'                            Chg : Edge, Chrome
' 2021/09/30  v0.1.0x  ishi  Chg : ExecuteScript
'                                    GetAttribute, GetInnerHTML, GetOuterHTML
'                                    SelectByIndex , DeSelectByIndex
'                                    Focus , FocusNoScroll, ScrollByElement
' 2021/10/16  v0.1.0y  ishi  Add : IE
'                            Chg : Start, CrateOptions, addOption
' 2021/10/24  v0.1.0z  ishi  Version is synced to WebElement.
'                            Chg : GetAttribute, GetInnerHTML, GetOuterHTML
'                                  SelectByIndex , DeSelectByIndex
'                            Del : getAttribute_js
' 2021/11/14  v0.1.1a  ishi  Add : DefualtPath
'                            Chg : Edge, Chrome, IE, Start
' 2021/11/23  v0.1.1b  ishi  Add : IsAlertPresent
' 2022/01/02  v0.1.1c0 ishi  Chg : CurrentURL, ExecuteScript, Execute
'                                    (Based on an issue #13,24,25,46 on GitHub)
'                                  IsElementPresent
'                                    (Based on an issue #27 on GitHub)
'                                  ImplicitlyWait -> SetImplicitlyWait
'                                    (ReName issue #24 on GitHub)
'                                  SetImplicitlyWait, SetPageLoadTimeout, SetScriptTimeout
'                                    (Based on an issue #24 on GitHub)
'                                  Changed non-Object or Variant arguments to explicit ByVal
'                                    (Double, Long, Boolean, String, etc.)
' 2022/01/02  v0.1.1c1 ishi  Chg : CurrentURL, ExecuteScript, Execute
'                                    (Based on an issue #46 on GitHub)
' 2022/01/06  v0.1.1c2 ishi  Chg : Bug Fix -> SelectByIndex, DeSelectByIndex
'                            Add : SelectAll
'                            Chg : DeSelectAll
' 2022/01/08  v0.1.1c3 ishi  Chg : Fixed a bug that sessionID is not specified.
'                                    CurrentURL
'                                    SelectAll, SelectByValue, SelectByVisibleText
'                                    SelectByIndex, SelectTagsOptions, AllSelectedOptions, SelectedOptionText
'                                    DeSelectAll, DeSelectByValue, DeSelectByVisibleText, DeSelectByIndex
'                                    Focus, FocusNoScroll, ScrollByElement
' 2022/01/10  v0.1.1c4 ishi  Add : HasAttribute
'                            Chg : Start , GetAttribute, SwitchToFrame, SelectByValue, SelectByVisibleText
'                                  AllSelectedOptions , SelectedOptionText, DeSelectByValue, DeSelectByVisibleText
'                                  ExecuteScript
'                            chg : Function ReName.
'                                    IsPresent, IsSelected
' ==========================================================================


Option Explicit

Public DefaultSessionId As String
Public UrlBase As String

' Driver commands
Private CMD_STATUS
Private CMD_NEW_SESSION
Private CMD_GET_ALL_SESSIONS
Private CMD_QUIT
Private CMD_GET_CURRENT_WINDOW_HANDLE
Private CMD_W3C_GET_CURRENT_WINDOW_HANDLE
Private CMD_GET_WINDOW_HANDLES
Private CMD_W3C_GET_WINDOW_HANDLES
Private CMD_GET
Private CMD_GO_FORWARD
Private CMD_GO_BACK
Private CMD_REFRESH
Private CMD_EXECUTE_SCRIPT
Private CMD_W3C_EXECUTE_SCRIPT
Private CMD_W3C_EXECUTE_SCRIPT_ASYNC
Private CMD_GET_CURRENT_URL
Private CMD_GET_TITLE
Private CMD_GET_PAGE_SOURCE
Private CMD_SCREENSHOT
Private CMD_ELEMENT_SCREENSHOT
Private CMD_FIND_ELEMENT
Private CMD_FIND_ELEMENTS
Private CMD_W3C_GET_ACTIVE_ELEMENT
Private CMD_GET_ACTIVE_ELEMENT
Private CMD_FIND_CHILD_ELEMENT
Private CMD_FIND_CHILD_ELEMENTS
Private CMD_CLICK_ELEMENT
Private CMD_CLEAR_ELEMENT
Private CMD_SUBMIT_ELEMENT
Private CMD_GET_ELEMENT_TEXT
Private CMD_SEND_KEYS_TO_ELEMENT
Private CMD_SEND_KEYS_TO_ACTIVE_ELEMENT
Private CMD_UPLOAD_FILE
Private CMD_GET_ELEMENT_VALUE
Private CMD_GET_ELEMENT_TAG_NAME
Private CMD_IS_ELEMENT_SELECTED
Private CMD_SET_ELEMENT_SELECTED
Private CMD_IS_ELEMENT_ENABLED
Private CMD_IS_ELEMENT_DISPLAYED
Private CMD_GET_ELEMENT_LOCATION
Private CMD_GET_ELEMENT_LOCATION_ONCE_SCROLLED_INTO_VIEW
Private CMD_GET_ELEMENT_SIZE
Private CMD_GET_ELEMENT_RECT
Private CMD_GET_ELEMENT_ATTRIBUTE
Private CMD_GET_ELEMENT_PROPERTY
Private CMD_GET_ALL_COOKIES
Private CMD_ADD_COOKIE
Private CMD_GET_COOKIE
Private CMD_DELETE_ALL_COOKIES
Private CMD_DELETE_COOKIE
Private CMD_SWITCH_TO_FRAME
Private CMD_SWITCH_TO_PARENT_FRAME
Private CMD_SWITCH_TO_WINDOW
Private CMD_CLOSE
Private CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY
Private CMD_IMPLICIT_WAIT
Private CMD_EXECUTE_ASYNC_SCRIPT
Private CMD_SET_SCRIPT_TIMEOUT
Private CMD_SET_TIMEOUTS
Private CMD_DISMISS_ALERT
Private CMD_W3C_DISMISS_ALERT
Private CMD_ACCEPT_ALERT
Private CMD_W3C_ACCEPT_ALERT
Private CMD_SET_ALERT_VALUE
Private CMD_W3C_SET_ALERT_VALUE
Private CMD_GET_ALERT_TEXT
Private CMD_W3C_GET_ALERT_TEXT
Private CMD_SET_ALERT_CREDENTIALS
Private CMD_CLICK
Private CMD_W3C_ACTIONS
Private CMD_W3C_CLEAR_ACTIONS
Private CMD_DOUBLE_CLICK
Private CMD_MOUSE_DOWN
Private CMD_MOUSE_UP
Private CMD_MOVE_TO
Private CMD_GET_WINDOW_SIZE
Private CMD_SET_WINDOW_SIZE
Private CMD_GET_WINDOW_POSITION
Private CMD_SET_WINDOW_POSITION
Private CMD_SET_WINDOW_RECT
Private CMD_GET_WINDOW_RECT
Private CMD_MAXIMIZE_WINDOW
Private CMD_W3C_MAXIMIZE_WINDOW
Private CMD_SET_SCREEN_ORIENTATION
Private CMD_GET_SCREEN_ORIENTATION
Private CMD_SINGLE_TAP
Private CMD_TOUCH_DOWN
Private CMD_TOUCH_UP
Private CMD_TOUCH_MOVE
Private CMD_TOUCH_SCROLL
Private CMD_DOUBLE_TAP
Private CMD_LONG_PRESS
Private CMD_FLICK
Private CMD_EXECUTE_SQL
Private CMD_GET_LOCATION
Private CMD_SET_LOCATION
Private CMD_GET_APP_CACHE
Private CMD_GET_APP_CACHE_STATUS
Private CMD_CLEAR_APP_CACHE
Private CMD_GET_NETWORK_CONNECTION
Private CMD_SET_NETWORK_CONNECTION
Private CMD_GET_LOCAL_STORAGE_ITEM
Private CMD_REMOVE_LOCAL_STORAGE_ITEM
Private CMD_GET_LOCAL_STORAGE_KEYS
Private CMD_SET_LOCAL_STORAGE_ITEM
Private CMD_CLEAR_LOCAL_STORAGE
Private CMD_GET_LOCAL_STORAGE_SIZE
Private CMD_GET_SESSION_STORAGE_ITEM
Private CMD_REMOVE_SESSION_STORAGE_ITEM
Private CMD_GET_SESSION_STORAGE_KEYS
Private CMD_SET_SESSION_STORAGE_ITEM
Private CMD_CLEAR_SESSION_STORAGE
Private CMD_GET_SESSION_STORAGE_SIZE
Private CMD_GET_LOG
Private CMD_GET_AVAILABLE_LOG_TYPES
Private CMD_CURRENT_CONTEXT_HANDLE
Private CMD_CONTEXT_HANDLES
Private CMD_SWITCH_TO_CONTEXT
Private CMD_FULLSCREEN_WINDOW
Private CMD_MINIMIZE_WINDOW
Private CMD_SHUTDOWN

Private CMD_GET_TIMEOUTS         '2021/7/14 add ishi


Private Const ELEMENT_KEY = "element-6066-11e4-a52e-4f735466cecf"

Public Enum By
    ID = 0
    TagName = 1
    ClassName = 2
    Name = 3
    CssSelector = 4
    XPath = 5                   '2021/7/15 add ishi
    LinkText = 6                '2021/7/15 add ishi
    PartialLinkText = 7         '2021/7/15 add ishi
End Enum

Public Enum AppWinStyle         '2021/8/10 add ishi
    vbHide = 0
    'vbNormalFocus = 1
    'vbMinimizedFocus = 2
    'vbMaximizedFocus = 3
    'vbNormalNoFocus = 4
    vbMinimizedNoFocus = 6
End Enum

'Set of special keys codes.     '2021/8/13 add ishi
Public Enum Keys
    NULL_key = &H0&                 '\ue000
    CANCEL_key = &H1&               '\ue001
    HELP_key = &H2&                 '\ue002
    BACKSPACE_key = &H3&            '\ue003
    BACK_SPACE_key = BACKSPACE_key
    TAB_key = &H4&                  '\ue004
    CLEAR_key = &H5&                '\ue005
    RETURN_key = &H6&               '\ue006
    ENTER_key = &H7&                '\ue007
    SHIFT_key = &H8&                '\ue008
    LEFT_SHIFT_key = SHIFT_key
    CONTROL_key = &H9&              '\ue009
    LEFT_CONTROL_key = CONTROL_key
    ALT_key = &HA&                  '\ue00a
    LEFT_ALT_key = ALT_key
    PAUSE_key = &HB&                '\ue00b
    ESCAPE_key = &HC&               '\ue00c
    SPACE_key = &HD&                '\ue00d
    PAGE_UP_key = &HE&              '\ue00e
    PAGE_DOWN_key = &HF&            '\ue00f
    END_key = &H10&                 '\ue010
    HOME_key = &H11&                '\ue011
    LEFT_key = &H12&                '\ue012
    ARROW_LEFT_key = LEFT_key
    UP_key = &H13&                  '\ue013
    ARROW_UP_key = UP_key
    RIGHT_key = &H14&               '\ue014
    ARROW_RIGHT_key = RIGHT_key
    DOWN_key = &H15&                '\ue015
    ARROW_DOWN_key = DOWN_key
    INSERT_key = &H16&              '\ue016
    DELETE_key = &H17&              '\ue017
    SEMICOLON_key = &H18&           '\ue018
    EQUALS_key = &H19&              '\ue019
    'number pad keys
    NUMPAD0_key = &H1A&             '\ue01a
    NUMPAD1_key = &H1B&             '\ue01b
    NUMPAD2_key = &H1C&             '\ue01c
    NUMPAD3_key = &H1D&             '\ue01d
    NUMPAD4_key = &H1E&             '\ue01e
    NUMPAD5_key = &H1F&             '\ue01f
    NUMPAD6_key = &H20&             '\ue020
    NUMPAD7_key = &H21&             '\ue021
    NUMPAD8_key = &H22&             '\ue022
    NUMPAD9_key = &H23&             '\ue023
    MULTIPLY_key = &H24&            '\ue024
    ADD_key = &H25&                 '\ue025
    SEPARATOR_key = &H26&           '\ue026
    SUBTRACT_key = &H27&            '\ue027
    DECIMAL_key = &H28&             '\ue028
    DIVIDE_key = &H29&              '\ue029
    'function  keys
    F1_key = &H31&                  '\ue031
    F2_key = &H32&                  '\ue032
    F3_key = &H33&                  '\ue033
    F4_key = &H34&                  '\ue034
    F5_key = &H35&                  '\ue035
    F6_key = &H36&                  '\ue036
    F7_key = &H37&                  '\ue037
    F8_key = &H38&                  '\ue038
    F9_key = &H39&                  '\ue039
    F10_key = &H3A&                 '\ue03a
    F11_key = &H3B&                 '\ue03b
    F12_key = &H3C&                 '\ue03c

    META_key = &H3D&                '\ue03d
    COMMAND_key = &H3D&             '\ue03d

    [_Min0] = NULL_key - 1
    [_Max0] = DIVIDE_key + 1
    [_Min1] = F1_key - 1
    [_Max1] = COMMAND_key + 1
End Enum

' ==========================================================================
' Win32 API
' ==========================================================================
' Sleep                         '2021/7/4 add ishi
#If VBA7 And Win64 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal millisecond As Long)
#End If

' GetSystemMetrics              '2021/7/13 add ishi
Private Const SM_CXSCREEN = 0           'Screen Withd (DPI)
Private Const SM_CYSCREEN = 1           'Screen Hright (DPI)
Private Const SM_CMOUSEBUTTONS = 43     'Mouse Bottons
Private Const SM_CMONITORS = 80         'Monitors
#If VBA7 And Win64 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If


' ==========================================================================
' Browser Options                       '2021/9/23 add ishi
' ==========================================================================
' DriverName
Private DriverName As String

' Browser specific value
Private browserName_data As String
Private platformName_data As String
Private platform_data As String
Private browserOptions_key As String

' Capabilities, DesiredCapabilities elements
Private extensions As Variant
Private args As Variant
Private Options As New Dictionary
Private alwaysMatch As New Dictionary
Private firstMatchDic As New Dictionary
Private firstMatch As Variant

' Capabilities, DesiredCapabilities
Private MyCapabilities As New Dictionary
Private MyDesiredCapabilities As New Dictionary


' ==========================================================================
' Browser Options
' ==========================================================================

' Browser specific value set            '2021/9/23 add ishi -> 2021/10/16 chg ishi
Public Sub CrateOptions()
    If DriverName = vbNullString Then
        Err.raise 1, "WebDriver.CrateOptions. DriverName is NullString."
    End If
    
    Select Case DriverName
    Case "msedgedriver"
        browserName_data = "MicrosoftEdge"
        platformName_data = "windows"
        platform_data = "WINDOWS"
        browserOptions_key = "ms:edgeOptions"
    Case "chromedriver"
        browserName_data = "chrome"
        platformName_data = "any"
        platform_data = "ANY"
        browserOptions_key = "goog:chromeOptions"
    Case "IEDriverServer"
        browserName_data = "internet explorer"
        platformName_data = "windows"
        platform_data = "windows"
        browserOptions_key = "se:ieOptions"
    Case Else
        Err.raise 2, "WebDriver.CrateOptions. DriverName is Not supported. (" & DriverName & ")"
    End Select
    
    ' Initialize to an empty array
    extensions = Array()
    args = Array()
End Sub

' Add Argument                          '2021/9/23 add ishi
Public Sub addArgument(ByVal argumentdata As String)
    If IsArray(args) Then
        ReDim Preserve args(UBound(args) + 1)
        args(UBound(args)) = argumentdata
    Else
        Err.raise 3, "WebDriver.addArgument. CrateOptions has not been called."
    End If
End Sub

'Add Option                             '2021/9/23 add ishi -> 2021/10/16 chg ishi
Public Sub addOption(ByVal Key As String, ByVal Data As Variant)
    If Options.Exists(Key) = False Then
        Options.Add Key, Data
    End If
End Sub

' Build Capabilities                    '2021/9/23 add ishi
Public Property Get BuildCapabilities() As Dictionary
    Options("extensions") = extensions
    Options("args") = args
    
    alwaysMatch("browserName") = browserName_data
    alwaysMatch("platformName") = platformName_data
    If alwaysMatch.Exists(browserOptions_key) = False Then
        alwaysMatch.Add browserOptions_key, Options
    End If
    
    firstMatch = Array(firstMatchDic)
    
    MyCapabilities("firstMatch") = firstMatch
    MyCapabilities.Add "alwaysMatch", alwaysMatch
    
    Set BuildCapabilities = MyCapabilities
End Property

' Build DesiredCapabilities             '2021/9/23 add ishi
Public Property Get BuildDesiredCapabilities() As Dictionary
    Options("extensions") = extensions
    Options("args") = args
   
    MyDesiredCapabilities("browserName") = browserName_data
    MyDesiredCapabilities("version") = ""
    MyDesiredCapabilities("platform") = platform_data
    If MyDesiredCapabilities.Exists(browserOptions_key) = False Then
        MyDesiredCapabilities.Add browserOptions_key, Options
    End If
    
    Set BuildDesiredCapabilities = MyDesiredCapabilities
End Property


' ==========================================================================
' Setup and shutdown
' ==========================================================================

' As WebDriver Default Path     '2021/11/14 add ishi
Private Property Get DefualtPath()
    DefualtPath = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\WebDriver\"    '?D????X??????
End Property

' Launch Edge Driver            '2021/8/10 chg ishi -> 2021/11/14 chg ishi
Public Sub Edge(Optional ByVal DriverPath As String = "msedgedriver.exe", _
                Optional ByVal driverUrl As String = "http://localhost:9515", _
                Optional ByVal VbAppWinStyle As AppWinStyle = vbMinimizedNoFocus)
    
    If DriverPath = vbNullString Then DriverPath = DefualtPath & "\msedgedriver.exe"
    If InStr(DriverPath, "\") = 0 Then DriverPath = DefualtPath & DriverPath
    Start DriverPath, driverUrl, VbAppWinStyle
    DriverName = "msedgedriver"
End Sub

' Launch Chrome Driver          '2021/8/10 chg ishi -> 2021/11/14 chg ishi
Public Sub Chrome(Optional ByVal DriverPath As String = "chromedriver.exe", _
                  Optional ByVal driverUrl As String = "http://localhost:9515", _
                  Optional ByVal VbAppWinStyle As AppWinStyle = vbMinimizedNoFocus)
    
    If DriverPath = vbNullString Then DriverPath = DefualtPath & "\chromedriver.exe"
    If InStr(DriverPath, "\") = 0 Then DriverPath = DefualtPath & DriverPath
    Start DriverPath, driverUrl, VbAppWinStyle
    DriverName = "chromedriver"
End Sub

' Launch IE Driver              '2021/10/16 add ishi -> 2021/11/14 chg ishi
Public Sub IE(Optional ByVal DriverPath As String = "IEDriverServer.exe", _
              Optional ByVal driverUrl As String = "http://localhost:5555", _
              Optional ByVal VbAppWinStyle As AppWinStyle = vbMinimizedNoFocus)
    
    If DriverPath = vbNullString Then DriverPath = DefualtPath & "\IEDriverServer.exe"
    If InStr(DriverPath, "\") = 0 Then DriverPath = DefualtPath & DriverPath
    Start DriverPath, driverUrl, VbAppWinStyle
    DriverName = "IEDriverServer"
End Sub

' Start WebDriver               '2022/1/10 chg
Private Sub Start(ByVal DriverPath As String, _
                  ByVal driverUrl As String, _
                  ByVal VbAppWinStyle As Long)
   
    ' Port
    Dim Port As String
    Port = vbNullString
    If InStr(DriverPath, "msedgedriver") > 0 Then
        Port = " --port=" & Mid(driverUrl, InStrRev(driverUrl, ":") + 1)
    ElseIf InStr(DriverPath, "chromedriver") > 0 Then
        Port = " --port=" & Mid(driverUrl, InStrRev(driverUrl, ":") + 1)
    ElseIf InStr(DriverPath, "IEDriverServer") > 0 Then
        Port = " /port=" & Mid(driverUrl, InStrRev(driverUrl, ":") + 1)
    End If
    
    ' Start WebDriver executable
    If Left(DriverPath, 1) <> Chr(34) Then DriverPath = Chr(34) & DriverPath
    If Right(DriverPath, 1) <> Chr(34) Then DriverPath = DriverPath & Chr(34)
    If Shell(DriverPath & Port, VbAppWinStyle) = 0 Then
        MsgBox "Failed in starting WebDriver." & vbCrLf & _
               "WebDriverPath : " & DriverPath & Port, vbCritical + vbOKOnly
        End
    End If
    
    ' Set WebDriver url
    UrlBase = driverUrl
    
    ' Initialize driver commands
    InitCommands
End Sub

' Shutdown WebDriver
Public Sub Shutdown()
    Execute CMD_SHUTDOWN
End Sub


' ==========================================================================
' Browser operations
' ==========================================================================

' Open browser              '2022/1/3 chg
Public Function OpenBrowser(Optional Capabilities As Dictionary = Nothing, _
                            Optional DesiredCapabilities As Dictionary = Nothing, _
                            Optional ByVal useAsDefault As Boolean = True) As String
    If Capabilities Is Nothing Then
        Set Capabilities = New Dictionary
    End If
    
    Dim Resp As Dictionary
    Set Resp = Execute(CMD_NEW_SESSION, Params("capabilities", Capabilities, _
                                               "desiredCapabilities", DesiredCapabilities))("value")
    
    If useAsDefault Then
        DefaultSessionId = Resp("sessionId")
    End If
    
    OpenBrowser = Resp("sessionId")
End Function

' Close browser
Public Sub CloseBrowser(Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Execute CMD_QUIT, Data
    DefaultSessionId = vbNullString
End Sub

' Open url                  '2022/1/3 chg
Public Function Navigate(ByVal url As String, _
                         Optional ByVal sessionId As String = vbNullString, _
                         Optional ByVal timeOutms, _
                         Optional ByVal raise As Boolean = True) As Boolean
    Dim Data As New Dictionary
    Dim Resp As Object
    Dim savtimeOutms
    
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Data.Add "url", url
    If Not IsMissing(timeOutms) Then 'GCUser99 added
        savtimeOutms = GetPageLoadTimeout(sessionId)
        If savtimeOutms <> timeOutms Then SetPageLoadTimeout timeOutms, sessionId
    End If
    
    Set Resp = Execute(CMD_GET, Data, raise)

    If Not IsMissing(timeOutms) Then  'GCUser99 added
        If savtimeOutms <> timeOutms Then SetPageLoadTimeout savtimeOutms, sessionId
    End If
    
    If IsResponseError(Resp) Then
        Navigate = False
        Debug.Print GetResponseErrorMessage(Resp)
    Else
        Navigate = True
    End If
End Function

' Current URL                   '2022/1/8 chg ishi
Public Function CurrentURL(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Dim encodeURL As String
    encodeURL = Execute(CMD_GET_CURRENT_URL, Data)("value")
    CurrentURL = ExecuteScript("return decodeURI('" & encodeURL & "')", , , sessionId)
End Function

' Windows Title                 '2022/1/3 chg
Public Function Title(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Title = Execute(CMD_GET_TITLE, Data)("value")
End Function

' Page Source                   '2022/1/3 chg
Public Function PageSource(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    PageSource = Execute(CMD_GET_PAGE_SOURCE, Data)("value")
End Function

' Switch To Next Window         '2021/6/23 add ishi
Public Function SwitchToNextWindow(Optional ByVal sessionId As String = vbNullString)
    Dim handles()   As String
    handles = GetWindowHandles
    SwitchToWindow handles(UBound(handles))
End Function

' Get Window Handles            '2022/1/3 chg
Public Function GetWindowHandles(Optional ByVal sessionId As String = vbNullString) As String()
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Dim handles
    Set handles = Execute(CMD_W3C_GET_WINDOW_HANDLES, Data)("value")
    
    ' To array of handles
    Dim Ret() As String
    Dim i As Integer
    For i = 0 To handles.Count - 1  ' handles is Collection, not array
        ReDim Preserve Ret(i)
        Ret(i) = handles.Item(i + 1)
    Next
    
    ' Return handles
    GetWindowHandles = Ret
End Function

' Switch To Window              '2021/6/23 add ishi
Public Function SwitchToWindow(ByVal handle As String, _
                               Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    Data.Add "handle", handle
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_SWITCH_TO_WINDOW, Data
End Function

' Get Current Window Handle     '2022/1/3 chg
Public Function GetCurrentWindowHandle(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    GetCurrentWindowHandle = Execute(CMD_W3C_GET_CURRENT_WINDOW_HANDLE, Data)("value")
End Function

' Close Window                  '2021/6/29 add ishi
Public Sub CloseWindow(Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Execute CMD_CLOSE, Data
End Sub

' Get Window Rect (px)          '2022/1/3 chg
Public Function GetWindowRect(Optional ByVal sessionId As String = vbNullString) As Integer()
    Dim Data As New Dictionary
    Dim rects
    Dim Ret() As Integer

    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Set rects = Execute(CMD_GET_WINDOW_RECT, Data)("value")
    
    ' To array of rects
    ReDim Preserve Ret(rects.Count - 1)
    Ret(0) = CInt(rects.Item("x"))
    Ret(1) = CInt(rects.Item("y"))
    Ret(2) = CInt(rects.Item("width"))
    Ret(3) = CInt(rects.Item("height"))
    
    ' Return rects
    GetWindowRect = Ret
End Function

' Set Window Rect (px)          '2022/1/3 chg
Public Function SetWindowRect(Optional ByVal x As Integer = 0, _
                              Optional ByVal y As Integer = 0, _
                              Optional ByVal width As Integer = 0, _
                              Optional ByVal height As Integer = 0, _
                              Optional ByVal sessionId As String = vbNullString) As Integer()
    Dim Data As New Dictionary
    Data.Add "x", x
    Data.Add "y", y
    Data.Add "width", width
    Data.Add "height", height
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Dim rects
    Set rects = Execute(CMD_SET_WINDOW_RECT, Data)("value")
    
    ' To array of rects
    Dim Ret() As Integer
    ReDim Preserve Ret(rects.Count - 1)
    Ret(0) = CInt(rects.Item("x"))
    Ret(1) = CInt(rects.Item("y"))
    Ret(2) = CInt(rects.Item("width"))
    Ret(3) = CInt(rects.Item("height"))
    
    ' Return rects
    SetWindowRect = Ret
End Function

' Get Window Size (px)          '2021/7/10 add ishi
Public Function GetWindowSize(Optional ByVal sessionId As String = vbNullString) As Integer()
    Dim rect()  As Integer
    rect = GetWindowRect(sessionId)
    
    Dim Ret(1)   As Integer
    Ret(0) = rect(2)    'width
    Ret(1) = rect(3)    'height
    
    ' Return
    GetWindowSize = Ret
End Function

' Set Window Size (px)          '2021/7/10 add ishi
Public Function SetWindowSize(ByVal width As Variant, _
                              ByVal height As Variant, _
                              Optional ByVal sessionId As String = vbNullString)
    Dim rect()  As Integer
    rect = GetWindowRect
    rect = SetWindowRect(rect(0), rect(1), width, height, sessionId)
End Function

' Get Window Position (px)      '2021/7/10 add ishi
Public Function GetWindowPosition(Optional ByVal sessionId As String = vbNullString) As Integer()
    Dim rect()  As Integer
    rect = GetWindowRect(sessionId)
    
    Dim Ret(1)   As Integer
    Ret(0) = rect(0)    'x
    Ret(1) = rect(1)    'y
    
    ' Return
    GetWindowPosition = Ret
End Function

' Set Window Position (px)      '2021/7/10 add ishi
Public Function SetWindowPosition(ByVal x As Variant, _
                                  ByVal y As Variant, _
                                  Optional ByVal sessionId As String = vbNullString)
    Dim rect()  As Integer
    rect = GetWindowRect
    rect = SetWindowRect(x, y, rect(2), rect(3), sessionId)
End Function

' Get Screen Size (dpi)         '2021/7/13 add ishi
Public Function GetScreenSize() As Integer()
    Dim Ret(1)   As Integer
    Ret(0) = GetSystemMetrics(SM_CXSCREEN)  'x
    Ret(1) = GetSystemMetrics(SM_CYSCREEN)  'y
    
    ' Return
    GetScreenSize = Ret
End Function

' Maximize Window               '2021/7/14 add ishi
Public Sub MaximizeWindow(Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Execute CMD_W3C_MAXIMIZE_WINDOW, Data
End Sub

' FullScreen Window             '2021/7/14 add ishi
Public Sub FullScreenWindow(Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Execute CMD_FULLSCREEN_WINDOW, Data
End Sub

' Minimize Window               '2021/7/14 add ishi
Public Sub MinimizeWindow(Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Execute CMD_MINIMIZE_WINDOW, Data
End Sub

' Save Screenshot               '2022/1/3 chg
Public Sub SaveScreenshot(ByVal filename As String, _
                          Optional ByVal ElementId As String = vbNullString, _
                          Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Dim png As String
    If ElementId = vbNullString Then
        png = Execute(CMD_SCREENSHOT, Data)("value")
    Else
        Data.Add "id", ElementId
        png = Execute(CMD_ELEMENT_SCREENSHOT, Data)("value")
    End If
    
    ' Conversion from base64 string to Byte array
    Dim chars() As Byte
    With CreateObject("MSXML2.DOMDocument").createElement("b64")
        .DataType = "bin.base64"
        .text = png
        chars = .nodeTypedValue
    End With
    
    ' png file put
    Dim fp As Long
    fp = FreeFile
    Open filename For Output As fp
    Close fp
    Open filename For Binary Access Write As fp
        Put #fp, 1, chars
    Close fp
End Sub


' ==========================================================================
' DOM operations
' ==========================================================================

' Find DOM element              '2022/1/3 chg
Public Function FindElement(by_ As By, _
                            ByVal value As String, _
                            Optional ByVal parentElementId As String = vbNullString, _
                            Optional ByVal sessionId As String = vbNullString) As WebElement
    Dim Data As Dictionary
    Set Data = ToSelector(by_, value)
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Dim cmd
    If parentElementId <> vbNullString Then
        Data.Add "id", parentElementId
        cmd = CMD_FIND_CHILD_ELEMENT
    Else
        cmd = CMD_FIND_ELEMENT
    End If
    
    ' Return element
    Set FindElement = ToWebElement(Execute(cmd, Data)("value")(ELEMENT_KEY), sessionId)
End Function

' Find multiple DOM elements    '2022/1/3 chg
Public Function FindElements(by_ As By, _
                             ByVal value As String, _
                             Optional ByVal parentElementId As String = vbNullString, _
                             Optional ByVal sessionId As String = vbNullString) As WebElement()
    Dim Data As Dictionary
    Set Data = ToSelector(by_, value)
    
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    
    Dim cmd
    If parentElementId <> vbNullString Then
        Data.Add "id", parentElementId
        cmd = CMD_FIND_CHILD_ELEMENTS
    Else
        cmd = CMD_FIND_ELEMENTS
    End If
    
    Dim elements
    Set elements = Execute(cmd, Data)("value")
    
    ' To array of ids
    Dim Ret() As WebElement
    Dim i As Integer
    For i = 0 To elements.Count - 1  ' elements is Collection, not array
        ReDim Preserve Ret(i)
        Set Ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY), sessionId)
    Next
    
    ' Return element ids
    FindElements = Ret
End Function

' Get Active Element            '2022/1/3 chg
Public Function GetActivElement(Optional ByVal sessionId As String = vbNullString) As WebElement
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
   ' Return element
   Set GetActivElement = ToWebElement(Execute(CMD_W3C_GET_ACTIVE_ELEMENT, Data)("value")(ELEMENT_KEY), sessionId)
End Function

' by* to CSS selector           '2021/7/15 chg ishi
Private Function ToSelector(by_ As By, _
                            ByVal value As String) As Dictionary
    Dim Data As New Dictionary
    Select Case by_
    Case By.ID
        Data.Add "using", "css selector"
        Data.Add "value", "[id=""" + value + """]"
    Case By.TagName
        Data.Add "using", "css selector"
        Data.Add "value", value
    Case By.ClassName
        Data.Add "using", "css selector"
        Data.Add "value", "." + value
    Case By.Name
        Data.Add "using", "css selector"
        Data.Add "value", "[name=""" + value + """]"
    Case By.CssSelector
        Data.Add "using", "css selector"
        Data.Add "value", value
    Case By.XPath
        Data.Add "using", "xpath"
        Data.Add "value", value
    Case By.LinkText
        Data.Add "using", "link text"
        Data.Add "value", value
    Case By.PartialLinkText
        Data.Add "using", "partial link text"
        Data.Add "value", value
    Case Else
        Data.Add "using", "css selector"
        Data.Add "value", value
    End Select
    
    Set ToSelector = Data
End Function

' Create element
Private Function ToWebElement(ByVal ElementId As String, _
                              Optional ByVal sessionId As String = vbNullString) As WebElement
    Dim element As New WebElement
    Set element.Driver_ = Me
    If sessionId = vbNullString Then
        element.SessionId_ = sessionId
    Else
        element.SessionId_ = DefaultSessionId
    End If
    element.ElementId_ = ElementId
    Set ToWebElement = element
End Function

' Returns element.value             '2022/1/3 chg
Public Function GetValue(ByVal ElementId As String, _
                         Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId
    Data.Add "name", "value"

    GetValue = Execute(CMD_GET_ELEMENT_ATTRIBUTE, Data)("value")
End Function

' Set value to element              '2021/8/13 chg ishi
' ( special keys codes support )
Public Sub SetValue(ByVal ElementId As String, _
                    ByVal text As Variant, _
                    Optional ClearBeforeTyping As Boolean = True, _
                    Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId
    
    Dim i As Integer
    Dim KeyCode As String
    Dim value As String
    Dim chars() As String
    If TypeName(text) = "Long" Then
        If Keys.[_Min0] < CLng(text) And CLng(text) < Keys.[_Max0] Or _
           Keys.[_Min1] < CLng(text) And CLng(text) < Keys.[_Max1] Then
            KeyCode = "\ue0" & Right("0" & LCase(Hex(CLng(text))), 2)
            value = KeyCode
            ReDim Preserve chars(0)
            chars(0) = KeyCode
        Else
            value = CStr(text)
            ReDim chars(Len(text) - 1)
            For i = 0 To UBound(chars)
                chars(i) = Mid(text, i + 1, 1)
            Next
        End If
    Else
        value = CStr(text)
        ReDim chars(Len(text) - 1)
        For i = 0 To UBound(chars)
            chars(i) = Mid(text, i + 1, 1)
        Next
    End If
    
    Data.Add "text", value
    Data.Add "value", chars
    
    If ClearBeforeTyping = True Then Execute CMD_CLEAR_ELEMENT, Data
    Execute CMD_SEND_KEYS_TO_ELEMENT, Data
End Sub

' Click element
Public Sub Click(ByVal ElementId As String, _
                 Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId

    Execute CMD_CLICK_ELEMENT, Data
End Sub

' Get text                      '2022/1/3 chg
Public Function GetText(ByVal ElementId As String, _
                        Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId

    GetText = Execute(CMD_GET_ELEMENT_TEXT, Data)("value")
End Function

' Get Property                  '2022/1/3 chg
Public Function GetProperty(ByVal value As String, _
                            ByVal ElementId As String, _
                            Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "name", value
    Data.Add "id", ElementId

    GetProperty = Execute(CMD_GET_ELEMENT_PROPERTY, Data)("value")
End Function

' Get Attribute                 '2021/10/24 chg ishi
'Public Function GetAttribute(ByVal value As String, _
'                             ByVal ElementId As String, _
'                             Optional ByVal sessionId As String = vbNullString) As String
'    Dim Script As String
'    If value = "innerHTML" Then
'        Script = "return arguments[0].innerHTML"
'    ElseIf value = "outerHTML" Then
'        Script = "return arguments[0].outerHTML"
'    Else
'        Script = "return arguments[0].getAttribute('" & value & "')"
'    End If
'
'    GetAttribute = ExecuteScript(Script, vbNullString, ElementId, sessionId)
'End Function

' Get Attribute                 '2022/1/10 chg CGuser99 -> chg ishi
Public Function GetAttribute(ByVal value As String, _
                             ByVal ElementId As String, _
                             Optional ByVal sessionId As String = vbNullString) As String
    Dim Script As String
    If value = "innerHTML" Then
        Script = "return arguments[0].innerHTML"
    ElseIf value = "outerHTML" Then
        Script = "return arguments[0].outerHTML"
    Else
        Script = "return arguments[0].getAttribute('" & value & "')"
    End If
    If HasAttribute(value, ElementId, sessionId) Then
        'below will return a VBA Null if attribute does not exist
        GetAttribute = ExecuteScript(Script, vbNullString, ElementId, sessionId)
    Else
        GetAttribute = ""
    End If
End Function

' Has Attribute                 '2022/1/10 add GCuser99 -> chg ishi
Public Function HasAttribute(ByVal value As String, _
                             ByVal ElementId As String, _
                             Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim Script As String
    Script = "return arguments[0].hasAttribute('" & value & "')"
    HasAttribute = ExecuteScript(Script, vbNullString, ElementId, sessionId)
End Function

' Get innerHTML                 '2021/10/24 chg ishi
Public Function GetInnerHTML(ByVal ElementId As String, _
                             Optional ByVal sessionId As String = vbNullString) As String
    Dim Script As String
    Script = "return arguments[0].innerHTML"
    
    GetInnerHTML = ExecuteScript(Script, vbNullString, ElementId, sessionId)
End Function

' Get outerHTML                 '2021/10/24 chg ishi
Public Function GetOuterHTML(ByVal ElementId As String, _
                             Optional ByVal sessionId As String = vbNullString) As String
    Dim Script As String
    Script = "return arguments[0].outerHTML"
    
    GetOuterHTML = ExecuteScript(Script, vbNullString, ElementId, sessionId)
End Function

' Clear value to element        '2021/7/4 add ishi
Public Sub Clear(ByVal ElementId As String, _
                 Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId
    Data.Add "text", vbNullString
    Data.Add "value", vbNullString
    
    Execute CMD_CLEAR_ELEMENT, Data
End Sub

' Is Present                    '2022/1/10 chg ishi
Public Function IsPresent(by_ As By, _
                          ByVal value As String, _
                          Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim timeout As Double
    Dim Resp As Dictionary
    Dim Data As Dictionary
    
    timeout = GetImplicitlyWait(sessionId)
    If timeout <> 0 Then SetImplicitlyWait 0, sessionId      ' ImplicitlyWait Disable
    
    Set Data = ToSelector(by_, value)
    If sessionId = vbNullString Then
        Data.Add "sessionId", DefaultSessionId
    Else
        Data.Add "sessionId", sessionId
    End If
    
    Set Resp = Execute(CMD_FIND_ELEMENT, Data, False)("value")  ' don't raise an error if not present

    If Not Resp.Exists("error") Then
        IsPresent = True
    Else
        IsPresent = False
    End If
    
    If timeout <> 0 Then SetImplicitlyWait timeout, sessionId
End Function

' Is Enabled                    '2022/1/3 chg
Public Function IsEnabled(ByVal ElementId As String, _
                          Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId
    
    IsEnabled = Execute(CMD_IS_ELEMENT_ENABLED, Data)("value")
End Function

' Is Displayed                  '2022/1/3 chg
Public Function IsDisplayed(ByVal ElementId As String, _
                            Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", ElementId

    IsDisplayed = Execute(CMD_IS_ELEMENT_DISPLAYED, Data)("value")
End Function


' ==========================================================================
' Frame operations
' ==========================================================================

' Switch to Frame               '2022/1/10 chg ishi
Public Function SwitchToFrame(ByVal ElementId As String, _
                              Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    Dim subdata As New Dictionary
    subdata.Add "ELEMENT", ElementId
    subdata.Add ELEMENT_KEY, ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Data.Add "id", subdata

    Execute CMD_SWITCH_TO_FRAME, Data
End Function

' Switch to Parent Frame        '2021/6/17 add ishi
Public Function SwitchToParentFrame(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_SWITCH_TO_PARENT_FRAME, Data
End Function

' Get Current Frame             '2022/1/3 chg
Public Function GetCurrentFrame(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    Data.Add "script", "return self.name"
    Dim args As Variant
    args = Array()
    Data.Add "args", args
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    GetCurrentFrame = Execute(CMD_W3C_EXECUTE_SCRIPT, Data)("value")
End Function


' ==========================================================================
' Alert operations
' ==========================================================================

' Get Alert Text                '2022/1/3 chg
Public Function GetAlertText(Optional ByVal sessionId As String = vbNullString) As String
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    GetAlertText = Execute(CMD_W3C_GET_ALERT_TEXT, Data)("value")
End Function

' Accept Alert                  '2021/6/20 add ishi
Public Function AcceptAlert(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_W3C_ACCEPT_ALERT, Data
End Function

' Dismiss Accept                '2021/6/20 add ishi
Public Function DismissAccept(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_W3C_DISMISS_ALERT, Data
End Function

' Is Alert Present              '2021/11/23 add ishi
Public Function IsAlertPresent(Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    On Error GoTo Err010
    Execute CMD_W3C_GET_ALERT_TEXT, Data
    IsAlertPresent = True
    Exit Function
    
Err010:
    IsAlertPresent = False
End Function


' ==========================================================================
' Select operations
' ==========================================================================

' Select All                    '2022/1/8 chg ishi
Public Function SelectAll(ByVal ElementId As String, _
                          Optional ByVal sessionId As String = vbNullString)
    Dim elmIds() As WebElement
    elmIds = FindChildElements(ElementId, sessionId)
    Dim i As Integer
    Focus ElementId, sessionId
    For i = 0 To UBound(elmIds)
        ExecuteScript "arguments[0].options[" & CStr(i) & "].selected = true;", vbNullString, ElementId, sessionId
    Next i
End Function

' Select By Value               '2022/1/10 chg ishi
Public Function SelectByValue(ByVal value As String, _
                              ByVal ElementId As String, _
                              Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    Data.Add "using", "css selector"
    Data.Add "value", "option[value =" & """" & value & """" & "]"
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Dim elmId As WebElement
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, Data)("value")(ELEMENT_KEY), sessionId)
    If Not IsSelected(elmId.ElementId_, sessionId) Then
        elmId.Click
    End If
End Function

' Select By Visible Text        '2022/1/10 chg ishi
Public Function SelectByVisibleText(ByVal text As String, _
                                    ByVal ElementId As String, _
                                    Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary
    Data.Add "using", "xpath"
    Data.Add "value", ".//option[normalize-space(.) = " & """" & text & """" & "]"
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Dim elmId As WebElement
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, Data)("value")(ELEMENT_KEY), sessionId)
    If Not IsSelected(elmId.ElementId_, sessionId) Then
        elmId.Click
    End If
End Function

' Select By Index               '2022/1/8 chg ishi
Public Function SelectByIndex(index As Integer, _
                              ByVal ElementId As String, _
                              Optional ByVal sessionId As String = vbNullString)
    Focus ElementId, sessionId
    ExecuteScript "arguments[0].options[" & CStr(index) & "].selected = true;", vbNullString, ElementId, sessionId
End Function

' Select Tag's Options          '2022/1/8 chg ishi
Public Function SelectTagsOptions(ByVal ElementId As String, _
                                  Optional ByVal sessionId As String = vbNullString) As String()
    Dim elmTags() As WebElement
    elmTags = FindElements(By.TagName, "option", ElementId, sessionId)
    Dim Ret() As String
    Dim i As Integer
    Ret = VBA.Split(VBA.vbNullString)    'Create Empty Array
    For i = 0 To UBound(elmTags)
        ReDim Preserve Ret(i)
        Ret(i) = elmTags(i).GetText
    Next
    SelectTagsOptions = Ret
End Function

' All Selected Options          '2022/1/10 chg ishi
Public Function AllSelectedOptions(ByVal ElementId As String, _
                                   Optional ByVal sessionId As String = vbNullString) As String()
    Dim elmIds() As WebElement
    elmIds = FindChildElements(ElementId, sessionId)
    Dim Ret() As String
    Dim i As Integer
    Dim j As Integer
    j = 0
    Ret = VBA.Split(VBA.vbNullString)    'Create Empty Array
    For i = 0 To UBound(elmIds)
        If IsSelected(elmIds(i).ElementId_, sessionId) Then
            ReDim Preserve Ret(j)
            Ret(j) = elmIds(i).GetText
            j = j + 1
        End If
    Next i
    AllSelectedOptions = Ret
End Function

' First Selected Option Text    '2022/1/10 chg ishi
Public Function SelectedOptionText(ByVal ElementId As String, _
                                   Optional ByVal sessionId As String = vbNullString) As String
    Dim elmIds() As WebElement
    elmIds = FindChildElements(ElementId, sessionId)
    Dim i As Integer
    For i = 0 To UBound(elmIds)
        If IsSelected(elmIds(i).ElementId_, sessionId) Then
            SelectedOptionText = elmIds(i).GetText
            Exit Function
        End If
    Next i
    SelectedOptionText = ""
End Function

' DeSelect All                  '2022/1/8 Chg ishi
Public Function DeSelectAll(ByVal ElementId As String, _
                            Optional ByVal sessionId As String = vbNullString)
    Dim elmIds() As WebElement
    elmIds = FindChildElements(ElementId, sessionId)
    Focus ElementId, sessionId
    Dim i As Integer
    For i = 0 To UBound(elmIds)
        ExecuteScript "arguments[0].options[" & CStr(i) & "].selected = false;", vbNullString, ElementId, sessionId
    Next i
End Function

' DeSelect By Value             '2022/1/10 chg ishi
Public Function DeSelectByValue(ByVal value As String, _
                                ByVal ElementId As String, _
                                Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary, elmId As WebElement
    Data.Add "using", "css selector"
    Data.Add "value", "option[value =" & """" & value & """" & "]"
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, Data)("value")(ELEMENT_KEY), sessionId)
    If IsSelected(elmId.ElementId_, sessionId) Then
        elmId.Click
    End If
End Function

' DeSelect By Visible Text      '2022/1/10 chg ishi
Public Function DeSelectByVisibleText(ByVal text As String, _
                                      ByVal ElementId As String, _
                                      Optional ByVal sessionId As String = vbNullString)
    Dim Data As New Dictionary, elmId As WebElement
    Data.Add "using", "xpath"
    Data.Add "value", ".//option[normalize-space(.) = " & """" & text & """" & "]"
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    Set elmId = ToWebElement(Execute(CMD_FIND_CHILD_ELEMENT, Data)("value")(ELEMENT_KEY), sessionId)
    If IsSelected(elmId.ElementId_, sessionId) Then
        elmId.Click
    End If
End Function

' DeSelect By Index             '2022/1/8 chg ishi
Public Function DeSelectByIndex(index As Integer, _
                                ByVal ElementId As String, _
                                Optional ByVal sessionId As String = vbNullString)
    Focus ElementId, sessionId
    ExecuteScript "arguments[0].options[" & CStr(index) & "].selected = false;", vbNullString, ElementId, sessionId
End Function

' Is Selected                   '2022/1/10 chg
Public Function IsSelected(ByVal ElementId As String, _
                           Optional ByVal sessionId As String = vbNullString) As Boolean
    Dim Data As New Dictionary
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If
    If Execute(CMD_IS_ELEMENT_SELECTED, Data)("value") = "True" Then
        IsSelected = True
    Else
        IsSelected = False
    End If
End Function

' Fine Child Elements           '2022/1/3 chg
Private Function FindChildElements(ByVal ElementId As String, _
                                   Optional ByVal sessionId As String = vbNullString) As WebElement()
    Dim Data As New Dictionary
    Dim elements
    Dim Ret() As WebElement
    Dim i As Integer
    
    Data.Add "using", "css selector"
    Data.Add "value", "option"
    Data.Add "id", ElementId
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Set elements = Execute(CMD_FIND_CHILD_ELEMENTS, Data)("value")
    
    ' To array
    For i = 0 To elements.Count - 1         'elements is Collection, not array
        ReDim Preserve Ret(i)
        Set Ret(i) = ToWebElement(elements(i + 1)(ELEMENT_KEY), sessionId)
    Next
    
    ' Return element ids
    FindChildElements = Ret
End Function


' ==========================================================================
' Wait operations
' ==========================================================================

' Set Implicitly Wait (millisecond)     '2022/1/10 chg
Public Function SetImplicitlyWait(Optional ByVal millisecond As Double = 0, _
                                  Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    Data.Add "implicit", millisecond
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_SET_TIMEOUTS, Data
End Function

' Set Page Load Timeout (millisecond)   '2022/1/2 chg
Public Function SetPageLoadTimeout(Optional ByVal millisecond As Double = 300000, _
                                   Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    Data.Add "pageLoad", millisecond
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_SET_TIMEOUTS, Data
End Function

' Set Script Timeout (millisecond)      '2022/1/2 chg
Public Function SetScriptTimeout(Optional ByVal millisecond As Double = 30000, _
                                 Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    Data.Add "script", millisecond
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_SET_TIMEOUTS, Data
End Function

' Get Implicitly Wait (millisecond)     '2022/1/3 chg
Public Function GetImplicitlyWait(Optional ByVal sessionId As String = vbNullString) As Double
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, Data)("value")

    GetImplicitlyWait = results.Item("implicit")
End Function

' Get Page Load Timeout (millisecond)   '2022/1/3 chg
Public Function GetPageLoadTimeout(Optional ByVal sessionId As String = vbNullString) As Double
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, Data)("value")

    GetPageLoadTimeout = results.Item("pageLoad")
End Function

' Get Script Timeout (millisecond)      '2022/1/3 chg
Public Function GetScriptTimeout(Optional ByVal sessionId As String = vbNullString) As Double
    Dim Data As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Dim results
    Set results = Execute(CMD_GET_TIMEOUTS, Data)("value")

    GetScriptTimeout = results.Item("script")
End Function

' Wait (millisecond)            '2021/7/4 add ishi
Public Function Wait(Optional millisecond As Long = 1000)
    Sleep millisecond
End Function


' ==========================================================================
' Navigation
' ==========================================================================

' Back                          '2021/8/7 add ishi
Public Function Back(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_GO_BACK, Data
End Function

' Forward                       '2021/8/7 add ishi
Public Function Forward(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_GO_FORWARD, Data
End Function

' Refresh                       '2021/8/7 add ishi
Public Function Refresh(Optional ByVal sessionId As String = vbNullString)
    Dim Data    As New Dictionary
    If sessionId <> vbNullString Then
        Data.Add "sessionId", sessionId
    End If

    Execute CMD_REFRESH, Data
End Function


' ==========================================================================
' Focus operations
' ==========================================================================

' Set Focus                     '2022/1/8 chg ishi
Public Function Focus(ByVal ElementId As String, _
                      Optional ByVal sessionId As String = vbNullString)
    Dim Script  As String
    Script = "arguments[0].focus({'preventScroll': arguments[1]})"
    
    ExecuteScript Script, 0, ElementId, sessionId   'Scroll
End Function

' Set Focus No Scroll           '2022/1/8 chg ishi
Public Function FocusNoScroll(ByVal ElementId As String, _
                              Optional ByVal sessionId As String = vbNullString)
    Dim Script  As String
    Script = "arguments[0].focus({'preventScroll': arguments[1]})"
    
    ExecuteScript Script, 1, ElementId, sessionId   'No Scroll
End Function


' ==========================================================================
' Scroll operations
' ==========================================================================

' Scroll By Element             '2022/1/8 chg ishi
Public Function ScrollByElement(ByVal ElementId As String, _
                                Optional ByVal Offset As Integer = 0, _
                                Optional ByVal sessionId As String = vbNullString)
    Dim Script As String
    Script = "arguments[0].scrollIntoView();"
    ExecuteScript Script, ElementId, ElementId, sessionId
    
    If Offset <> 0 Then
        Script = "window.scrollTo(0, window.pageYOffset + " & CStr(Offset) & ");"
        ExecuteScript Script, ElementId, ElementId, sessionId
    End If
End Function


' ==========================================================================
' Common functions
' ==========================================================================

' Execute Script                '2022/1/10 chg
Public Function ExecuteScript(ByVal Script As String, _
                              Optional ScriptArgs As Variant = vbNullString, _
                              Optional ByVal ElementId As String = vbNullString, _
                              Optional ByVal sessionId As String = vbNullString, _
                              Optional ByVal timeOutms, _
                              Optional ByVal raise As Boolean = True)
    Dim Data As New Dictionary
    Dim ElmData As New Dictionary
    Dim args As Variant
    Dim savtimeOutms

    Data.Add "script", Script
    
    ' Set ElementID and Args
    If ElementId <> vbNullString Then
        ElmData.Add "ELEMENT", ElementId
        ElmData.Add ELEMENT_KEY, ElementId
        args = Array(ElmData)
        If VarType(ScriptArgs) = vbVariant + vbArray Then    'Scriptargs is Array
            Dim i As Integer
            For i = 0 To UBound(ScriptArgs)
                ReDim Preserve args(i + 1)
                args(i + 1) = ScriptArgs(i)
            Next i
        Else
            args = Array(ElmData, ScriptArgs)
        End If
    Else
        args = Array()
    End If
    Data.Add "args", args
    
    ' Set sessionId
    If sessionId = vbNullString Then
        Data.Add "sessionId", DefaultSessionId
    Else
        Data.Add "sessionId", sessionId
    End If
    
    If Not IsMissing(timeOutms) Then
        savtimeOutms = GetScriptTimeout(sessionId)
        If savtimeOutms <> timeOutms Then SetScriptTimeout timeOutms, sessionId
    End If
    
    Dim Resp As Dictionary
    Set Resp = Execute(CMD_W3C_EXECUTE_SCRIPT, Data, raise)
    
    'trap error if occurs   2022/1/3 ishi
    If IsResponseError(Resp) Then
        Err.raise 513, "WebDriver.ExecuteScript", GetResponseErrorMessage(Resp)
    Else
        ExecuteScript = Resp("value")
        If Not IsMissing(timeOutms) Then
            If savtimeOutms <> timeOutms Then SetScriptTimeout savtimeOutms, sessionId
        End If
    End If
End Function

' Execute driver command            '2022/1/3 chg
Public Function Execute(driverCommand, _
                        Optional parameters As Dictionary = Nothing, _
                        Optional ByVal raise As Boolean = True) As Dictionary
    Dim method As String: method = driverCommand(0)
    Dim Path As String: Path = driverCommand(1)
    Dim Resp As Dictionary
    'Dim JsonConverter As New JsonConverter
    
    If parameters Is Nothing Then
        Set parameters = New Dictionary
    End If
    
    ' Set default session id if session id is missing
    If Not parameters.Exists("sessionId") Then
        parameters.Add "sessionId", DefaultSessionId
    End If
    
    Dim paramKey As Variant
    Dim nonPathParameters As New Dictionary
    For Each paramKey In parameters
        'VarType(parameters(paramKey)) = vbString And
        If InStr(Path, "$" & paramKey) > 0 Then 'path parameter
            Path = Replace(Path, "$" & paramKey, parameters(paramKey))
        Else 'non-path parameter
            nonPathParameters.Add paramKey, parameters(paramKey)
        End If
    Next paramKey
    
    'Send request to selenium server
    Set Resp = SendRequest(method, UrlBase + Path, nonPathParameters)
    
    If raise And IsResponseError(Resp) Then 'trap error if occurs
        Err.raise 513, "WebDriver.Execute", GetResponseErrorMessage(Resp)
    End If

    Set Execute = Resp 'always pass a dictionary object - let caller parse info based on context, including error
    
End Function

Private Function GetResponseErrorMessage(Resp As Dictionary) As String
    GetResponseErrorMessage = ""
    If TypeName(Resp("value")) = "Dictionary" Then
        If Resp("value").Exists("error") Then
            GetResponseErrorMessage = JsonConverter.ConvertToJson(Resp("value"))
        End If
    End If
End Function

Private Function IsResponseError(Resp As Dictionary) As Boolean
    IsResponseError = False
    If TypeName(Resp("value")) = "Dictionary" Then
        If Resp("value").Exists("error") Then
            IsResponseError = True
        End If
    End If
End Function

' Send HTTP request
Private Function SendRequest(ByVal method As String, _
                             ByVal url As String, _
                             Optional Data As Dictionary = Nothing) As Dictionary
    Dim client As Object
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    
    client.Open method, url
    
    If method = "POST" Or method = "PUT" Then
        client.setRequestHeader "Content-Type", "application/json"
        client.send JsonConverter.ConvertToJson(Data)
    Else
        client.send
    End If

    Do While client.readyState < 4
        DoEvents
    Loop

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(client.responseText)
    Set SendRequest = Json
End Function


' ==========================================================================
' Driver commands
' ==========================================================================
Private Sub InitCommands()
    CMD_STATUS = Array("GET", "/status")
    CMD_NEW_SESSION = Array("POST", "/session")
    CMD_GET_ALL_SESSIONS = Array("GET", "/sessions")
    CMD_QUIT = Array("DELETE", "/session/$sessionId")
    CMD_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window_handle")
    CMD_W3C_GET_CURRENT_WINDOW_HANDLE = Array("GET", "/session/$sessionId/window")
    CMD_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window_handles")
    CMD_W3C_GET_WINDOW_HANDLES = Array("GET", "/session/$sessionId/window/handles")
    CMD_GET = Array("POST", "/session/$sessionId/url")
    CMD_GO_FORWARD = Array("POST", "/session/$sessionId/forward")
    CMD_GO_BACK = Array("POST", "/session/$sessionId/back")
    CMD_REFRESH = Array("POST", "/session/$sessionId/refresh")
    CMD_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute")
    CMD_W3C_EXECUTE_SCRIPT = Array("POST", "/session/$sessionId/execute/sync")
    CMD_W3C_EXECUTE_SCRIPT_ASYNC = Array("POST", "/session/$sessionId/execute/async")
    CMD_GET_CURRENT_URL = Array("GET", "/session/$sessionId/url")
    CMD_GET_TITLE = Array("GET", "/session/$sessionId/title")
    CMD_GET_PAGE_SOURCE = Array("GET", "/session/$sessionId/source")
    CMD_SCREENSHOT = Array("GET", "/session/$sessionId/screenshot")
    CMD_ELEMENT_SCREENSHOT = Array("GET", "/session/$sessionId/element/$id/screenshot")
    CMD_FIND_ELEMENT = Array("POST", "/session/$sessionId/element")
    CMD_FIND_ELEMENTS = Array("POST", "/session/$sessionId/elements")
    CMD_W3C_GET_ACTIVE_ELEMENT = Array("GET", "/session/$sessionId/element/active")
    CMD_GET_ACTIVE_ELEMENT = Array("POST", "/session/$sessionId/element/active")
    CMD_FIND_CHILD_ELEMENT = Array("POST", "/session/$sessionId/element/$id/element")
    CMD_FIND_CHILD_ELEMENTS = Array("POST", "/session/$sessionId/element/$id/elements")
    CMD_CLICK_ELEMENT = Array("POST", "/session/$sessionId/element/$id/click")
    CMD_CLEAR_ELEMENT = Array("POST", "/session/$sessionId/element/$id/clear")
    CMD_SUBMIT_ELEMENT = Array("POST", "/session/$sessionId/element/$id/submit")
    CMD_GET_ELEMENT_TEXT = Array("GET", "/session/$sessionId/element/$id/text")
    CMD_SEND_KEYS_TO_ELEMENT = Array("POST", "/session/$sessionId/element/$id/value")
    CMD_SEND_KEYS_TO_ACTIVE_ELEMENT = Array("POST", "/session/$sessionId/keys")
    CMD_UPLOAD_FILE = Array("POST", "/session/$sessionId/file")
    CMD_GET_ELEMENT_VALUE = Array("GET", "/session/$sessionId/element/$id/value")
    CMD_GET_ELEMENT_TAG_NAME = Array("GET", "/session/$sessionId/element/$id/name")
    CMD_IS_ELEMENT_SELECTED = Array("GET", "/session/$sessionId/element/$id/selected")
    CMD_SET_ELEMENT_SELECTED = Array("POST", "/session/$sessionId/element/$id/selected")
    CMD_IS_ELEMENT_ENABLED = Array("GET", "/session/$sessionId/element/$id/enabled")
    CMD_IS_ELEMENT_DISPLAYED = Array("GET", "/session/$sessionId/element/$id/displayed")
    CMD_GET_ELEMENT_LOCATION = Array("GET", "/session/$sessionId/element/$id/location")
    CMD_GET_ELEMENT_LOCATION_ONCE_SCROLLED_INTO_VIEW = Array("GET", "/session/$sessionId/element/$id/location_in_view")
    CMD_GET_ELEMENT_SIZE = Array("GET", "/session/$sessionId/element/$id/size")
    CMD_GET_ELEMENT_RECT = Array("GET", "/session/$sessionId/element/$id/rect")
    CMD_GET_ELEMENT_ATTRIBUTE = Array("GET", "/session/$sessionId/element/$id/attribute/$name")
    CMD_GET_ELEMENT_PROPERTY = Array("GET", "/session/$sessionId/element/$id/property/$name")
    CMD_GET_ALL_COOKIES = Array("GET", "/session/$sessionId/cookie")
    CMD_ADD_COOKIE = Array("POST", "/session/$sessionId/cookie")
    CMD_GET_COOKIE = Array("GET", "/session/$sessionId/cookie/$name")
    CMD_DELETE_ALL_COOKIES = Array("DELETE", "/session/$sessionId/cookie")
    CMD_DELETE_COOKIE = Array("DELETE", "/session/$sessionId/cookie/$name")
    CMD_SWITCH_TO_FRAME = Array("POST", "/session/$sessionId/frame")
    CMD_SWITCH_TO_PARENT_FRAME = Array("POST", "/session/$sessionId/frame/parent")
    CMD_SWITCH_TO_WINDOW = Array("POST", "/session/$sessionId/window")
    CMD_CLOSE = Array("DELETE", "/session/$sessionId/window")
    CMD_GET_ELEMENT_VALUE_OF_CSS_PROPERTY = Array("GET", "/session/$sessionId/element/$id/css/$propertyName")
    CMD_IMPLICIT_WAIT = Array("POST", "/session/$sessionId/timeouts/implicit_wait")
    CMD_EXECUTE_ASYNC_SCRIPT = Array("POST", "/session/$sessionId/execute_async")
    CMD_SET_SCRIPT_TIMEOUT = Array("POST", "/session/$sessionId/timeouts/async_script")
    CMD_SET_TIMEOUTS = Array("POST", "/session/$sessionId/timeouts")
    CMD_DISMISS_ALERT = Array("POST", "/session/$sessionId/dismiss_alert")
    CMD_W3C_DISMISS_ALERT = Array("POST", "/session/$sessionId/alert/dismiss")
    CMD_ACCEPT_ALERT = Array("POST", "/session/$sessionId/accept_alert")
    CMD_W3C_ACCEPT_ALERT = Array("POST", "/session/$sessionId/alert/accept")
    CMD_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert_text")
    CMD_W3C_SET_ALERT_VALUE = Array("POST", "/session/$sessionId/alert/text")
    CMD_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert_text")
    CMD_W3C_GET_ALERT_TEXT = Array("GET", "/session/$sessionId/alert/text")
    CMD_SET_ALERT_CREDENTIALS = Array("POST", "/session/$sessionId/alert/credentials")
    CMD_CLICK = Array("POST", "/session/$sessionId/click")
    CMD_W3C_ACTIONS = Array("POST", "/session/$sessionId/actions")
    CMD_W3C_CLEAR_ACTIONS = Array("DELETE", "/session/$sessionId/actions")
    CMD_DOUBLE_CLICK = Array("POST", "/session/$sessionId/doubleclick")
    CMD_MOUSE_DOWN = Array("POST", "/session/$sessionId/buttondown")
    CMD_MOUSE_UP = Array("POST", "/session/$sessionId/buttonup")
    CMD_MOVE_TO = Array("POST", "/session/$sessionId/moveto")
    CMD_GET_WINDOW_SIZE = Array("GET", "/session/$sessionId/window/$windowHandle/size")
    CMD_SET_WINDOW_SIZE = Array("POST", "/session/$sessionId/window/$windowHandle/size")
    CMD_GET_WINDOW_POSITION = Array("GET", "/session/$sessionId/window/$windowHandle/position")
    CMD_SET_WINDOW_POSITION = Array("POST", "/session/$sessionId/window/$windowHandle/position")
    CMD_SET_WINDOW_RECT = Array("POST", "/session/$sessionId/window/rect")
    CMD_GET_WINDOW_RECT = Array("GET", "/session/$sessionId/window/rect")
    CMD_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/$windowHandle/maximize")
    CMD_W3C_MAXIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/maximize")
    CMD_SET_SCREEN_ORIENTATION = Array("POST", "/session/$sessionId/orientation")
    CMD_GET_SCREEN_ORIENTATION = Array("GET", "/session/$sessionId/orientation")
    CMD_SINGLE_TAP = Array("POST", "/session/$sessionId/touch/click")
    CMD_TOUCH_DOWN = Array("POST", "/session/$sessionId/touch/down")
    CMD_TOUCH_UP = Array("POST", "/session/$sessionId/touch/up")
    CMD_TOUCH_MOVE = Array("POST", "/session/$sessionId/touch/move")
    CMD_TOUCH_SCROLL = Array("POST", "/session/$sessionId/touch/scroll")
    CMD_DOUBLE_TAP = Array("POST", "/session/$sessionId/touch/doubleclick")
    CMD_LONG_PRESS = Array("POST", "/session/$sessionId/touch/longclick")
    CMD_FLICK = Array("POST", "/session/$sessionId/touch/flick")
    CMD_EXECUTE_SQL = Array("POST", "/session/$sessionId/execute_sql")
    CMD_GET_LOCATION = Array("GET", "/session/$sessionId/location")
    CMD_SET_LOCATION = Array("POST", "/session/$sessionId/location")
    CMD_GET_APP_CACHE = Array("GET", "/session/$sessionId/application_cache")
    CMD_GET_APP_CACHE_STATUS = Array("GET", "/session/$sessionId/application_cache/status")
    CMD_CLEAR_APP_CACHE = Array("DELETE", "/session/$sessionId/application_cache/clear")
    CMD_GET_NETWORK_CONNECTION = Array("GET", "/session/$sessionId/network_connection")
    CMD_SET_NETWORK_CONNECTION = Array("POST", "/session/$sessionId/network_connection")
    CMD_GET_LOCAL_STORAGE_ITEM = Array("GET", "/session/$sessionId/local_storage/key/$key")
    CMD_REMOVE_LOCAL_STORAGE_ITEM = Array("DELETE", "/session/$sessionId/local_storage/key/$key")
    CMD_GET_LOCAL_STORAGE_KEYS = Array("GET", "/session/$sessionId/local_storage")
    CMD_SET_LOCAL_STORAGE_ITEM = Array("POST", "/session/$sessionId/local_storage")
    CMD_CLEAR_LOCAL_STORAGE = Array("DELETE", "/session/$sessionId/local_storage")
    CMD_GET_LOCAL_STORAGE_SIZE = Array("GET", "/session/$sessionId/local_storage/size")
    CMD_GET_SESSION_STORAGE_ITEM = Array("GET", "/session/$sessionId/session_storage/key/$key")
    CMD_REMOVE_SESSION_STORAGE_ITEM = Array("DELETE", "/session/$sessionId/session_storage/key/$key")
    CMD_GET_SESSION_STORAGE_KEYS = Array("GET", "/session/$sessionId/session_storage")
    CMD_SET_SESSION_STORAGE_ITEM = Array("POST", "/session/$sessionId/session_storage")
    CMD_CLEAR_SESSION_STORAGE = Array("DELETE", "/session/$sessionId/session_storage")
    CMD_GET_SESSION_STORAGE_SIZE = Array("GET", "/session/$sessionId/session_storage/size")
    CMD_GET_LOG = Array("POST", "/session/$sessionId/log")
    CMD_GET_AVAILABLE_LOG_TYPES = Array("GET", "/session/$sessionId/log/types")
    CMD_CURRENT_CONTEXT_HANDLE = Array("GET", "/session/$sessionId/context")
    CMD_CONTEXT_HANDLES = Array("GET", "/session/$sessionId/contexts")
    CMD_SWITCH_TO_CONTEXT = Array("POST", "/session/$sessionId/context")
    CMD_FULLSCREEN_WINDOW = Array("POST", "/session/$sessionId/window/fullscreen")
    CMD_MINIMIZE_WINDOW = Array("POST", "/session/$sessionId/window/minimize")
    CMD_SHUTDOWN = Array("GET", "/shutdown")
    CMD_GET_TIMEOUTS = Array("GET", "/session/$sessionId/timeouts")         '2021/7/14 add ishi
End Sub


' ==========================================================================
' Utility functions
' ==========================================================================

' KeyValue arguments to parameters as dictionaly
Private Function Params(ParamArray keysAndValues()) As Dictionary
    Dim dict As New Dictionary
    Dim i As Integer
    For i = 0 To UBound(keysAndValues) - 1 Step 2
        dict.Add keysAndValues(i), keysAndValues(i + 1)
    Next i
    Set Params = dict
End Function


