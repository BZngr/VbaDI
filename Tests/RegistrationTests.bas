Attribute VB_Name = "RegistrationTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module

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

'@TestMethod("VbaDI.Container.Register")
Private Sub RegisteredWithEmptyParamArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
        
    'Act:
    xSut.Register

    'Assert:
    'The test is that an exception is not thrown
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Register")
Private Sub RegistrationCollectionBeforeSingle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(TS.CreateVbaDITestObject(1000), "5")
    
    Dim xRegs As Collection
    Set xRegs = New Collection
    
    Dim xV As Long
    For xV = 1 To 10
        xRegs.Add VbaDI.Instance(TS.CreateVbaDITestObject(xV), CStr(xV))
    Next
        
    'Act:
    xSut.Register xRegs, xReg

    'Assert:
    Dim xActualReg As VbaDIRegistration
    Dim xActual As VbaDITestObject
    Set xActualReg = RegistrationCode.GetRegistration( _
        DataProvider.GetAddInData().RegistrationsByObjectID, "5")
    Set xActual = xActualReg.Instance
    
    Assert.AreEqual "5", xActual.InstanceID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Register")
Private Sub RegistrationSingleBeforeCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(TS.CreateVbaDITestObject(1000), "5")
    
    Dim xRegs As Collection
    Set xRegs = New Collection
    
    Dim xV As Long
    For xV = 1 To 10
        xRegs.Add VbaDI.Instance(TS.CreateVbaDITestObject(xV), CStr(xV))
    Next
        
    'Act:
    xSut.Register xReg, xRegs

    Dim xActualReg As VbaDIRegistration
    Dim xActual As VbaDITestObject
    
    Set xActualReg = RegistrationCode.GetRegistration( _
        DataProvider.GetAddInData().RegistrationsByObjectID, "5")
    Set xActual = xActualReg.Instance
    
    'Assert:
    Assert.AreEqual "1000", xActual.InstanceID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Register")
Private Sub LifestyleMismatchAbstractErrorTest()
    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_LIFESTYLE_MISMATCH
    VbaDIError.EnablePrintToImmediateWindow False
    On Error GoTo TestFail

    'Arrange:
    Dim xConcrete As IVbaDIFluentRegistration
    Set xConcrete = VbaDI.Instance(New VbaDITestObject)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xConcrete


    Dim xAbstract As IVbaDIFluentRegistration
    Set xAbstract = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(New VbaDITestObject) _
        .AsTransient()
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xAbstract

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.Container.Register")
Private Sub LifestyleMismatchConcreteErrorTest()
    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_LIFESTYLE_MISMATCH
    VbaDIError.EnablePrintToImmediateWindow False
    On Error GoTo TestFail

    'Arrange:
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()

    Dim xAbstract As IVbaDIFluentRegistration
    Set xAbstract = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(New VbaDITestObject)

    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xAbstract

    Dim xConcrete As IVbaDIFluentRegistration
    Set xConcrete = VbaDI.Instance(New VbaDITestObject).AsTransient()
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xConcrete

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.Container.Register")
Private Sub LifestyleMismatchAbstractsErrorTest()
    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_LIFESTYLE_MISMATCH
    VbaDIError.EnablePrintToImmediateWindow False
    On Error GoTo TestFail

    'Arrange:
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()

    Dim xAbstract As IVbaDIFluentRegistration
    Set xAbstract = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(New VbaDITestObject)
    
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xAbstract
    
    Dim xAbstract2 As IVbaDIFluentRegistration
    Set xAbstract2 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface2)) _
        .Use(New VbaDITestObject).AsTransient
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xAbstract2

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("VbaDI.Container.Register")
Private Sub InvalidRegisterParameterParamErrorTest()
    Dim ExpectedError As Long
    ExpectedError = 5
    VbaDIError.EnablePrintToImmediateWindow False
    On Error GoTo TestFail

    'Arrange:
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()

    Dim xColl As Collection
    Set xColl = New Collection
    xColl.Add VbaDI.ForInterface(TypeName(New IVbaDITestInterface2)) _
        .Use(New VbaDITestObject).AsTransient
    xColl.Add 2
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    xSut.Register xColl

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("VbaDI.Container.Register")
Private Sub InvalidLoadParameterParamErrorTest()
    Dim ExpectedError As Long
    ExpectedError = 5
    VbaDIError.EnablePrintToImmediateWindow False
    On Error GoTo TestFail

    'Arrange:
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()

    Dim xColl As Collection
    Set xColl = New Collection
    xColl.Add New VbaDITestRegistrationLoader
    xColl.Add 2
    
    'Act:
    '@Ignore FunctionReturnValueDiscarded
    xSut.RegisterUsingLoader xColl

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


