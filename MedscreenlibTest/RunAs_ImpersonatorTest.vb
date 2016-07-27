﻿Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports MedscreenLib



'''<summary>
'''This is a test class for RunAs_ImpersonatorTest and is intended
'''to contain all RunAs_ImpersonatorTest Unit Tests
'''</summary>
<TestClass()> _
Public Class RunAs_ImpersonatorTest


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
    '''A test for ImpersonateStart
    '''</summary>
    <TestMethod()> _
    Public Sub ImpersonateStartTest()
        Dim target As RunAs_Impersonator = New RunAs_Impersonator ' TODO: Initialize to an appropriate value
        Dim Domain As String = "Concateno" ' TODO: Initialize to an appropriate value
        Dim userName As String = "svc.dartimage" ' TODO: Initialize to an appropriate value
        Dim Password As String = "D4RT1mag3" ' TODO: Initialize to an appropriate value
        target.ImpersonateStart(Domain, userName, Password)
        Assert.Inconclusive("A method that does not return a value cannot be verified.")
    End Sub
End Class
