Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports MedscreenLib



'''<summary>
'''This is a test class for MedscreenTest and is intended
'''to contain all MedscreenTest Unit Tests
'''</summary>
<TestClass()> _
Public Class MedscreenTest


    Private testContextInstance As TestContext

    '''<summary>
    '''Gets or sets the test context which provides
    '''information about and functionality for the current test run.
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "Additional test attributes"
    '
    'You can use the following additional attributes as you write your tests:
    '
    'Use ClassInitialize to run code before running the first test in the class
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    'Use ClassCleanup to run code after all tests in a class have run
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    'Use TestInitialize to run code before running each test
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    'Use TestCleanup to run code after each test has run
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    '''<summary>
    '''A test for GetLoginId
    '''</summary>
    <TestMethod()> _
    Public Sub GetLoginIdTest()
        Dim username As String = "David Lawson" ' TODO: Initialize to an appropriate value
        Dim expected As String = "crippsk" ' TODO: Initialize to an appropriate value
        Dim actual As String
        actual = Medscreen.GetLoginId(username)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for GetLoginId
    '''</summary>
    <TestMethod()> _
    Public Sub GetLoginIdTest1()
        Dim username As String = String.Empty ' TODO: Initialize to an appropriate value
        Dim expected As String = String.Empty ' TODO: Initialize to an appropriate value
        Dim actual As String
        actual = Medscreen.GetLoginId(username)
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for FindUsers
    '''</summary>
    <TestMethod()> _
    Public Sub FindUsersTest()
        Dim stub As String = "Lawson" ' TODO: Initialize to an appropriate value
        Dim expected As String = String.Empty ' TODO: Initialize to an appropriate value
        Dim actual As String
        actual = Medscreen.FindUsers(stub)
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for MatchWholeWord
    '''</summary>
    <TestMethod()> _
    Public Sub MatchWholeWordTest()
        Dim Source As String = " Confirm" ' TODO: Initialize to an appropriate value
        Dim target As String = "CONFIRM" ' TODO: Initialize to an appropriate value
        Dim expected As Integer = 1 ' TODO: Initialize to an appropriate value
        Dim actual As Integer
        actual = Medscreen.MatchWholeWord(Source, target)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub
    <TestMethod()> _
    Public Sub MatchWholeWordTest1()
        Dim Source As String = " Confirm " ' TODO: Initialize to an appropriate value
        Dim target As String = "CONFIRM" ' TODO: Initialize to an appropriate value
        Dim expected As Integer = 1 ' TODO: Initialize to an appropriate value
        Dim actual As Integer
        actual = Medscreen.MatchWholeWord(Source, target)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    <TestMethod()> _
    Public Sub MatchWholeWordTest2()
        Dim Source As String = " Confirm." ' TODO: Initialize to an appropriate value
        Dim target As String = "CONFIRM" ' TODO: Initialize to an appropriate value
        Dim expected As Integer = 1 ' TODO: Initialize to an appropriate value
        Dim actual As Integer
        actual = Medscreen.MatchWholeWord(Source, target)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub
    <TestMethod()> _
    Public Sub MatchWholeWordTest3()
        Dim Source As String = " Confirmed" ' TODO: Initialize to an appropriate value
        Dim target As String = "CONFIRM" ' TODO: Initialize to an appropriate value
        Dim expected As Integer = 0 ' TODO: Initialize to an appropriate value
        Dim actual As Integer
        actual = Medscreen.MatchWholeWord(Source, target)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for ReplaceLastCharWithString
    '''</summary>
    <TestMethod()> _
    Public Sub ReplaceLastCharWithStringTest()
        Dim inString As String = "one or two," ' TODO: Initialize to an appropriate value
        Dim inChar As Char = "," ' TODO: Initialize to an appropriate value
        Dim repStr As String = " and " ' TODO: Initialize to an appropriate value
        Dim expected As String = "one or two and " ' TODO: Initialize to an appropriate value
        Dim actual As String
        actual = Medscreen.ReplaceLastCharWithString(inString, inChar, repStr)
        Assert.AreEqual(expected, actual)
        'Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub

    '''<summary>
    '''A test for Capitalise
    '''</summary>
    <TestMethod()> _
    Public Sub CapitaliseTest()
        Dim strIn As String = "FRED BLOGGS LLP " ' TODO: Initialize to an appropriate value
        Dim expected As String = String.Empty ' TODO: Initialize to an appropriate value
        Dim actual As String
        actual = Medscreen.Capitalise(strIn)
        Assert.AreEqual(expected, actual)
        Assert.Inconclusive("Verify the correctness of this test method.")
    End Sub
End Class
