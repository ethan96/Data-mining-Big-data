Imports OpenQA.Selenium
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Support.UI
Imports System.Linq.Expressions

Public Class WebDriverUtil

    Public Shared Function GetParent(e As IWebElement) As IWebElement
        Return e.FindElement(By.XPath(".."))
    End Function

    Public Shared Function IsElementReady(FirefoxDriver1 As FirefoxDriver, Selector As OpenQA.Selenium.By, Optional TimeOutSec As Integer = 10) As Boolean
        For i As Integer = 0 To TimeOutSec - 1
            Try
                If FirefoxDriver1.FindElement(Selector) IsNot Nothing Then Return True
            Catch ex As InvalidElementStateException
                Threading.Thread.Sleep(999)
            Catch ex2 As NoSuchElementException
                Threading.Thread.Sleep(999)
            End Try
        Next
        Return False
    End Function

    Public Shared Sub SwitchToWindow(driver As FirefoxDriver, predicateExp As Expression(Of Func(Of IWebDriver, Boolean)))
        Dim predicate = predicateExp.Compile()
        For Each handle In driver.WindowHandles
            driver.SwitchTo().Window(handle)
            If predicate(driver) Then
                Return
            End If
        Next

        Throw New ArgumentException(String.Format("Unable to find window with condition: '{0}'", predicateExp.Body))
    End Sub

End Class
