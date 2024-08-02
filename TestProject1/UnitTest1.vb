Imports OpenQA.Selenium
Imports OpenQA.Selenium.Edge
Imports OpenQA.Selenium.Firefox
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Serilog
Imports FluentAssertions

' Enables parallel test execution at the method level
<Assembly: Parallelize(Workers:=6, Scope:=ExecutionScope.MethodLevel)>

Namespace Task
    <TestClass>
    Public Class UnitTest1
        Public Class LoginPage
            Private driver As IWebDriver
            Private logger As ILogger

            Public Sub New(driver As IWebDriver, logger As ILogger)
                Me.driver = driver
                Me.logger = logger
                driver.Navigate().GoToUrl("https://www.saucedemo.com/")
            End Sub

            Public ReadOnly Property UsernameField As IWebElement
                Get
                    Return driver.FindElement(By.CssSelector("#user-name"))
                End Get
            End Property

            Public ReadOnly Property PasswordField As IWebElement
                Get
                    Return driver.FindElement(By.CssSelector("#password"))
                End Get
            End Property

            Public ReadOnly Property LoginButton As IWebElement
                Get
                    Return driver.FindElement(By.CssSelector("#login-button"))
                End Get
            End Property

            Public ReadOnly Property ErrorMessage As IWebElement
                Get
                    Return driver.FindElement(By.CssSelector(".error-message-container"))
                End Get
            End Property

            Public ReadOnly Property Dashboard As IWebElement
                Get
                    Return driver.FindElement(By.CssSelector(".app_logo"))
                End Get
            End Property

            Public Sub Login(username As String, password As String)
                UsernameField.SendKeys(username)
                PasswordField.SendKeys(password)
                logger.Information("Login fields written")
            End Sub

            Public Sub Click()
                LoginButton.Click()
                logger.Information("Login button clicked")
            End Sub

            Public Sub ClearInputs(Optional user As Boolean = True, Optional password As Boolean = True)
                If user Then
                    UsernameField.SendKeys(Keys.Control + "a")
                    UsernameField.SendKeys(Keys.Delete)
                End If
                If password Then
                    PasswordField.SendKeys(Keys.Control + "a")
                    PasswordField.SendKeys(Keys.Delete)
                End If
                logger.Information("Login fields cleared")
            End Sub

            Public Function Validate(str As String) As Boolean
                Return Dashboard.Text.Equals(str)
            End Function
        End Class

        Public Shared Function CreateDriver(browser As String, logger As ILogger) As IWebDriver
            If browser.Equals("edge") Then
                logger.Information("Edge driver created")
                Return New EdgeDriver()
            ElseIf browser.Equals("firefox") Then
                logger.Information("Firefox driver created")
                Return New FirefoxDriver()
            Else
                logger.Error("Creation of driver failed!")
                Throw New Exception("Invalid browser specified")
            End If
        End Function

        Private logger As ILogger

        <TestInitialize>
        Public Sub Setup()
            logger = New LoggerConfiguration().WriteTo.File("logs.txt").WriteTo.Console().CreateLogger()
        End Sub

        <TestMethod>
        <DataRow("edge", "standard_user", "secret_sauce")>
        <DataRow("firefox", "visual_user", "secret_sauce")>
        Public Sub UC_1(browser As String, user As String, password As String)
            logger.Information("START OF UC_1")
            Dim driver As IWebDriver = CreateDriver(browser, logger)
            Dim loginPage As New LoginPage(driver, logger)
            loginPage.Login(user, password)
            loginPage.ClearInputs()
            loginPage.Click()
            loginPage.ErrorMessage.Text.Should().Contain("Username is required")
            driver.Dispose()
            logger.Information("END OF UC_1")
        End Sub

        <TestMethod>
        <DataRow("edge", "standard_user", "secret_sauce")>
        <DataRow("firefox", "visual_user", "secret_sauce")>
        Public Sub UC_2(browser As String, user As String, password As String)
            logger.Information("START OF UC_2")
            Dim driver As IWebDriver = CreateDriver(browser, logger)
            Dim loginPage As New LoginPage(driver, logger)
            loginPage.Login(user, password)
            loginPage.ClearInputs(False, True)
            loginPage.Click()
            loginPage.ErrorMessage.Text.Should().Contain("Password is required")
            driver.Dispose()
            logger.Information("END OF UC_2")
        End Sub

        <TestMethod>
        <DataRow("edge", "standard_user", "secret_sauce")>
        <DataRow("firefox", "visual_user", "secret_sauce")>
        Public Sub UC_3(browser As String, user As String, password As String)
            logger.Information("START OF UC_3")
            Dim driver As IWebDriver = CreateDriver(browser, logger)
            Dim loginPage As New LoginPage(driver, logger)
            loginPage.Login(user, password)
            loginPage.Click()
            loginPage.Validate("Swag Labs").Should().BeTrue()
            driver.Dispose()
            logger.Information("END OF UC_3")
        End Sub
    End Class
End Namespace