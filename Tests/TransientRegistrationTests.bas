Attribute VB_Name = "TransientRegistrationTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
'@Ignore VariableNotUsed
Private Fakes As Rubberduck.FakesProvider
Private TS As VbaDITestSupport

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
    Set TS = New VbaDITestSupport
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set TS = Nothing
End Sub

'@TestInitialize
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("VbaDI.Container.Resolve.Transient")
Private Sub TransientReturnsNewObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xTO1 As VbaDITestObject
    Set xTO1 = New VbaDITestObject
    xTO1.InstanceID = 1000

    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.Instance(xTO1).AsTransient()

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg1

    'Act:
    Dim xActual As VbaDITestObject
    Set xActual = xSut.Resolve(TypeName(New VbaDITestObject))

    'Assert:
    Assert.IsTrue TypeOf xActual Is VbaDITestObject, "xContainer.Resolve returned 'Nothing'"
    Assert.AreNotEqual xTO1.InstanceID, xActual.InstanceID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Transient")
Private Sub TransientReturnsNewResolvedObject()
    On Error GoTo TestFail

    'Arrange:
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.InstanceID = 1000
    xTO.AddValueDependency "TestValue"

    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.Instance(xTO) _
        .DependsOnValue("TestValue", "12345").AsTransient()

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg1


    Dim xContainer As VbaDIContainer
    Set xContainer = xSut
    xContainer.SetResolver TS.Resolver

    'Act:
    Dim xActual As VbaDITestObject
    Set xActual = xSut.Resolve(TypeName(New VbaDITestObject))

    'Assert:
    Assert.AreNotEqual xTO.InstanceID, xActual.InstanceID
    Assert.AreEqual "12345", xActual.InjectedValueDependency("TestValue")
    Assert.AreEqual CLng(0), xActual.GetDependencyIDCalls, "GetDependencyIDs not called"
    Assert.AreEqual CLng(1), xActual.InjectDependenciesCalls, "InjectDependencies not called"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VbaDI.Container.Resolve.Transient")
Private Sub TransientWithTransientDependency()
    On Error GoTo TestFail

    'Arrange:
    Dim xTO1 As VbaDITestObject
    Set xTO1 = New VbaDITestObject
    xTO1.InstanceID = 1000
    xTO1.AddValueDependency "TestValue"

    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.Instance(xTO1, xTO1.InstanceID) _
        .DependsOnValue("TestValue", "123435").AsTransient()

    Dim xTO2 As VbaDITestObject
    Set xTO2 = New VbaDITestObject
    xTO2.InstanceID = 2000
    xTO2.AddObjectIDDependency "1000"

    Dim xReg2 As IVbaDIFluentRegistration
    Set xReg2 = VbaDI.Instance(xTO2, xTO2.InstanceID).AsTransient()

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg1
    xSut.Register xReg2

    Dim xContainer As VbaDIContainer
    Set xContainer = xSut
    xContainer.SetResolver TS.Resolver

    'Act:

    Dim xActual2 As VbaDITestObject
    Set xActual2 = xSut.Resolve("2000")

    'Assert:
    Assert.AreNotEqual xTO2.InstanceID, xActual2.InstanceID

    Dim xObjects As Collection
    Set xObjects = xActual2.InjectedObjects()
    Assert.IsTrue xObjects.Count > 0

    Dim xActual1 As VbaDITestObject
On Error Resume Next
    Set xActual1 = xActual2.InjectedObject("1000")
On Error GoTo 0
    Assert.IsNotNothing xActual1
    Assert.AreNotEqual "1000", xActual1.InstanceID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


