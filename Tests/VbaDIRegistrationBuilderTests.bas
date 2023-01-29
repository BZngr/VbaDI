Attribute VB_Name = "VbaDIRegistrationBuilderTests"
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
    VbaDIError.EnablePrintToImmediateWindow True
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub AddInterfaceIDsArray()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As IVbaDIFluentRegistration
    Set xSut = VbaDI.Instance(New VbaDITestObject)
    
    Dim xExpected As Object
    Set xExpected = util.CreateDictionary()
    With xExpected
        .Add "One", 1
        .Add "Two", 2
        .Add "Three", 3
    End With
    
    'Act:
    
    Set xSut = xSut.ForInterface(Array("One", "Two", "Three"))
    Dim xReg As VbaDIRegistration
    Set xReg = RegistrationCode.CreateRegistration(xSut)
    'Assert:
    
    Dim xItf As Variant
    For Each xItf In xReg.InterfaceIDs
        Assert.IsTrue xExpected.Exists(xItf)
    Next
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub AddInterfaceIDsCollection()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As IVbaDIFluentRegistration
    Set xSut = VbaDI.Instance(New VbaDITestObject)
    
    Dim xExpected As Object
    Set xExpected = util.CreateDictionary()
    With xExpected
        .Add "One", 1
        .Add "Two", 2
        .Add "Three", 3
    End With
    Dim xIDs As Variant
    xIDs = xExpected.Keys()
    
    'Act:
    
    Set xSut = xSut.ForInterface(xIDs)
    
    'Assert:
    
    Dim xReg As VbaDIRegistration
    Set xReg = RegistrationCode.CreateRegistration(xSut)
    
    Dim xItf As Variant
    For Each xItf In xReg.InterfaceIDs
        Assert.IsTrue xExpected.Exists(xItf)
    Next
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub AddInterfaceNonCStrConvertibleInput()
    Dim ExpectedError As Long: ExpectedError = VbaDIError.ERROR_INVALID_REGISTRATION_ID
    On Error GoTo TestFail

    'Arrange:
    VbaDIError.EnablePrintToImmediateWindow False
    
    Dim xSut As IVbaDIFluentRegistration
    Set xSut = VbaDI.Instance(New VbaDITestObject)

    'Act:
    '@Ignore AssignmentNotUsed
    Set xSut = xSut.ForInterface(New VbaDITestObject)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub AddInterfaceCollectionOfObjects()
    Dim ExpectedError As Long: ExpectedError = VbaDIError.ERROR_INVALID_REGISTRATION_ID
    On Error GoTo TestFail

    'Arrange:
    VbaDIError.EnablePrintToImmediateWindow False
    Dim xInput As Collection
    Set xInput = New Collection
    With xInput
        .Add New VbaDITestObject
        .Add New VbaDITestObject
        .Add New VbaDITestObject
    End With
    Dim xSut As IVbaDIFluentRegistration
    Set xSut = VbaDI.Instance(New VbaDITestObject)

    'Act:
    '@Ignore AssignmentNotUsed
    Set xSut = xSut.ForInterface(xInput)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub ValueDependencySet()
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim xExpected As String
    xExpected = "ValueString"
    
    Dim xTestObj1 As VbaDITestObject
    Set xTestObj1 = TS.CreateVbaDITestObject(1)
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(xTestObj1) _
        .DependsOnValue("TestVal", xExpected)
        
    'Assert:
    Dim xActual As String
    xActual = RegistrationCode.CreateRegistration(xReg1).ValueDependencies.Item("TestVal")
    Assert.AreEqual xExpected, xActual
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub ValueDependencySetBeforeInstance()
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim xExpected As String
    xExpected = "ValueString"
    
    Dim xTestObj1 As VbaDITestObject
    Set xTestObj1 = TS.CreateVbaDITestObject(1)
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .DependsOnValue("TestVal", xExpected) _
        .Use(xTestObj1)
        
    'Assert:
    Dim xActual As String
    xActual = RegistrationCode.CreateRegistration(xReg1) _
        .ValueDependencies.Item("TestVal")
    Assert.AreEqual xExpected, xActual
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.VbaDIRegistrationBuilder")
Private Sub SetLifestyleBeforeInstance()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xTestObj1 As VbaDITestObject
    Set xTestObj1 = TS.CreateVbaDITestObject(1)
    
    'Act:
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI _
        .ForInterface(TypeName(New IVbaDITestInterface)) _
        .AsTransient() _
        .Use(xTestObj1)
                
    'Assert:
    Dim xActual As Boolean
    xActual = RegistrationCode.CreateRegistration(xReg1).IsSingleton
    Assert.AreEqual False, xActual
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Registration")
Private Sub RegistrationIncompleteJustLifetyleErrorTest()

    VbaDIError.EnablePrintToImmediateWindow False

    Dim ExpectedError As Long: ExpectedError = VbaDIError.ERROR_REGISTRATION_INCOMPLETE
    On Error GoTo TestFail

    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface("IGuess")
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()

    'Act:
    xContainer.Register xReg
    

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.Registration")
Private Sub RegistrationIncompleteJustValueErrorTest()

    VbaDIError.EnablePrintToImmediateWindow False

    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_REGISTRATION_INCOMPLETE
    On Error GoTo TestFail

    'Arrange:
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface("IGuess") _
        .DependsOnValue("TestVal", "Nope!")
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()

    'Act:
    xContainer.Register xReg

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


