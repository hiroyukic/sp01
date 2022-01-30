VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VERSION 1.0 CLASS

'BEGIN
 ' MultiUse = -1  'True
'End
'Attribute VB_Name = "WebElement"
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
' 2021/06/23  v0.1.0a  ishi  Add
' 2021/06/25  v0.1.0d  ishi  Add : Switch
' 2021/06/25  v0.1.0e  ishi  Add : GetAttribute
' 2021/06/28  v0.1.0f  ishi  Version is synced to WebDriver.
' 2021/06/30  v0.1.0g  ishi  Add : innerHTML, outerHTML
' 2021/07/04  v0.1.0h  ishi  Add : Clear
' 2021/07/09  v0.1.0i  ishi  Version is synced to WebDriver.
' 2021/07/10  v0.1.0j  ishi  Add : SetFocus, SetFocusNoScroll
' 2021/07/13  v0.1.0k  ishi  Add : IsEnabled, IsDisplayed
' 2021/07/14  v0.1.0l  ishi  Version is synced to WebDriver.
' 2021/07/15  v0.1.0m  ishi  Version is synced to WebDriver.
' 2021/07/16  v0.1.0n  ishi  Version is synced to WebDriver.
' 2021/07/20  v0.1.0o  ishi  Version is synced to WebDriver.
' 2021/07/24  v0.1.0p  ishi  Add : Scroll
' 2021/07/28  v0.1.0q  ishi  Version is synced to WebDriver.
' 2021/08/07  v0.1.0r  ishi  Add : GetProperty
' 2021/08/10  v0.1.0s  ishi  Add : SelectByVisibleText, SelectByIndex
'                            Add : SelectTagsOptions, AllSelectedOptions
'                            Add : DeSelectAll, DeSelectByValue
'                            Add : DeSelectByVisibleText, DeSelectByIndex
' 2021/08/13  v0.1.0t  ishi  Chg : SetValue
' 2021/08/14  v0.1.0u  ishi  Add : SaveScreen
' 2021/08/22  v0.1.0v  ishi  Version is synced to WebDriver.
' 2021/09/23  v0.1.0w  ishi  Version is synced to WebDriver.
' 2021/09/30  v0.1.0x  ishi  Add : ExecuteScript
' 2021/10/16  v0.1.0y  ishi  Version is synced to WebDriver.
' 2021/10/24  v0.1.0z  ishi  Chg : ExecuteScript
' 2021/11/14  v0.1.1a  ishi  Version is synced to WebDriver.
' 2021/11/23  v0.1.1b  ishi  Version is synced to WebDriver.
' 2022/01/02  v0.1.1c0 ishi  Version is synced to WebDriver.
' 2022/01/02  v0.1.1c1 ishi  Version is synced to WebDriver.
' 2022/01/06  v0.1.1c2 ishi  Version is synced to WebDriver.
'                            Add : SelectAll
' 2022/01/06  v0.1.1c3 ishi  Version is synced to WebDriver.
' 2022/01/10  v0.1.1c4 ishi  Add : HasAttribute, IsSelected
' ==========================================================================


Option Explicit

Public Driver_ As WebDriver
Public SessionId_ As String
Public ElementId_ As String


' ==========================================================================
' Browser operations
' ==========================================================================

' Save Screenshot                   '2021/8/14 add ishi
Public Sub SaveScreenshot(ByVal filename As String)
    Driver_.SaveScreenshot filename, ElementId_, SessionId_
End Sub


' ==========================================================================
' DOM operations
' ==========================================================================

' Find DOM element
Public Function FindElement(by_ As By, ByVal value As String) As WebElement
    Set FindElement = Driver_.FindElement(by_, value, ElementId_, SessionId_)
End Function

' Find multiple DOM elements
Public Function FindElements(by_ As By, ByVal value As String) As WebElement()
    FindElements = Driver_.FindElements(by_, value, ElementId_, SessionId_)
End Function

' Returns element.value
Public Function GetValue() As String
    GetValue = Driver_.GetValue(ElementId_, SessionId_)
End Function

' Set value to element              '2021/8/13 chg ishi
Public Sub SetValue(ByVal text As Variant, _
                    Optional ClearBeforeTyping As Boolean = True)
    Driver_.SetValue ElementId_, text, ClearBeforeTyping, SessionId_
End Sub

' Click
Public Sub Click()
    Driver_.Click ElementId_, SessionId_
End Sub

' Returns element.innerText
Public Function GetText() As String
    GetText = Driver_.GetText(ElementId_, SessionId_)
End Function

' Returns element.property          '2021/8/7 add ishi
Public Function GetProperty(ByVal value As String) As String
    GetProperty = Driver_.GetProperty(value, ElementId_, SessionId_)
End Function

' Returns element.GetAttribute      '2021/6/25 add ishi
Public Function GetAttribute(ByVal value As String) As String
    GetAttribute = Driver_.GetAttribute(value, ElementId_, SessionId_)
End Function

' Returns element.HasAttribute      '2022/1/10 add
Public Function HasAttribute(ByVal value As String) As String
    HasAttribute = Driver_.HasAttribute(value, ElementId_, SessionId_)
End Function

' Returns element.innerHTML         '2021/6/30 add ishi
Public Function GetInnerHTML() As String
    GetInnerHTML = Driver_.GetInnerHTML(ElementId_, SessionId_)
End Function

' Returns element.outerHTML         '2021/6/30 add ishi
Public Function GetOuterHTML() As String
    GetOuterHTML = Driver_.GetOuterHTML(ElementId_, SessionId_)
End Function

' Clear value to element            '2021/7/4 add ishi
Public Sub Clear()
    Driver_.Clear ElementId_, SessionId_
End Sub

' Set Focus                         '2021/7/10 add ishi
Public Sub Focus()
    Driver_.Focus ElementId_, SessionId_
End Sub

' Set Focus No Scroll               '2021/7/10 add ishi
Public Sub FocusNoScroll()
    Driver_.FocusNoScroll ElementId_, SessionId_
End Sub

' Is Enabled                        '2021/7/13 add ishi
Public Function IsEnabled() As Boolean
    IsEnabled = Driver_.IsEnabled(ElementId_, SessionId_)
End Function

' Is Displayed                      '2021/7/13 add ishi
Public Function IsDisplayed() As Boolean
    IsDisplayed = Driver_.IsDisplayed(ElementId_, SessionId_)
End Function


' ==========================================================================
' Frame operations
' ==========================================================================

' Switch to Frame                   '2021/6/22 add ishi
Public Function SwitchToFrame()
    Driver_.SwitchToFrame ElementId_, SessionId_
End Function


' ==========================================================================
' Select operations
' ==========================================================================

' Select All                        '2022/1/6 add ishi
Public Function SelectAll() As String()
    Driver_.SelectAll ElementId_, SessionId_
End Function

' Select By Value                   '2021/6/22 add ishi
Public Function SelectByValue(ByVal value As String)
    Driver_.SelectByValue value, ElementId_, SessionId_
End Function

' Select By Visible Text            '2021/8/10 add ishi
Public Function SelectByVisibleText(ByVal text As String)
    Driver_.SelectByVisibleText text, ElementId_, SessionId_
End Function

' Select By Index                   '2021/8/10 add ishi
Public Function SelectByIndex(ByVal index As Integer)
    Driver_.SelectByIndex index, ElementId_, SessionId_
End Function

' Select Tag's Options              '2021/8/10 add ishi
Public Function SelectTagsOptions() As String()
    SelectTagsOptions = Driver_.SelectTagsOptions(ElementId_, SessionId_)
End Function

' All Selected Options              '2021/8/10 add ishi
Public Function AllSelectedOptions() As String()
    AllSelectedOptions = Driver_.AllSelectedOptions(ElementId_, SessionId_)
End Function

' First Selected Option Text        '2021/6/22 add ishi
Public Function SelectedOptionText() As String
    SelectedOptionText = Driver_.SelectedOptionText(ElementId_, SessionId_)
End Function

' DeSelect All                      '2021/8/10 add ishi
Public Function DeSelectAll() As String()
    Driver_.DeSelectAll ElementId_, SessionId_
End Function

' DeSelect By Value                 '2021/8/10 add ishi
Public Function DeSelectByValue(ByVal value As String)
    Driver_.DeSelectByValue value, ElementId_, SessionId_
End Function

' DeSelect By Visible Text          '2021/8/10 add ishi
Public Function DeSelectByVisibleText(ByVal text As String)
    Driver_.DeSelectByVisibleText text, ElementId_, SessionId_
End Function

' DeSelect By Index                 '2021/8/10 add ishi
Public Function DeSelectByIndex(ByVal index As Integer)
    Driver_.DeSelectByIndex index, ElementId_, SessionId_
End Function

' IsSelected                        '2022/1/10 add ishi
Public Function IsSelected() As Boolean
    IsSelected = Driver_.IsSelected(ElementId_, SessionId_)
End Function


' ==========================================================================
' Scroll operations
' ==========================================================================

' Scroll By Element                 '2021/7/24 add ishi
Public Function ScrollByElement(Optional ByVal Offset As Integer = 0)
    Driver_.ScrollByElement ElementId_, Offset, SessionId_
End Function


' ==========================================================================
' ExecuteScript operations
' ==========================================================================

' ExecuteScript                     '2021/9/30 add ishi -> 2021/10/24 chg ishi
Public Function ExecuteScript(ByVal Script As String, _
                              Optional ScriptArgs As Variant = vbNullString)
    ExecuteScript = Driver_.ExecuteScript(Script, ScriptArgs, ElementId_, SessionId_)
End Function



