Attribute VB_Name = "VbaDIContainerTests"
'@IgnoreModule FunctionReturnValueDiscarded, EmptyMethod
'@TestModule
'@Folder("VbaDI.Tests")
Option Explicit
Option Private Module

Private Type TRegData
    Instance As Object
    ObjectID As Variant
End Type

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
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    VbaDIError.EnablePrintToImmediateWindow True
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub AssociatesInterfaceImplementation()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(New VbaDITestObject)
    
    xContainer.Register xReg
    
    'Act:
    Dim xSut As VbaDITestObject
    Set xSut = xContainer.Resolve(TypeName(New IVbaDITestInterface))

    'Assert:
    Assert.IsTrue Not xSut Is Nothing, "Object not resolved"
    Assert.IsTrue xSut.ReturnInputValue(5) = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub AssociatesInterfaceImplementationWithSingletonDependency()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    
    Dim xTestObj As VbaDITestObject
    Set xTestObj = New VbaDITestObject
    xTestObj.AddObjectIDDependency "Collection"
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(xTestObj)
    
    xContainer.Register xReg
    xContainer.Register VbaDI.Instance(New Collection, "Collection")
    
    'Act:
    Dim xSut As VbaDITestObject
    Set xSut = xContainer.Resolve(TypeName(New IVbaDITestInterface))

    'Assert:
    Assert.IsTrue Not xSut Is Nothing, "Object not resolved"
    Assert.IsTrue xSut.ReturnInputValue(5) = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub ReturnsObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(New VbaDITestObject)
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    
    xContainer.Register xReg
    
    'Act:
    Dim xSut As VbaDITestObject
    Set xSut = xContainer.Resolve(TypeName(New VbaDITestObject))

    'Assert:
    Assert.IsTrue Not xSut Is Nothing, "Object not resolved"
    Assert.IsTrue xSut.ReturnInputValue(5) = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Load")
Private Sub LoadMany()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xLoader As VbaDITestRegistrationLoader
    Set xLoader = New VbaDITestRegistrationLoader
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xLoaders As Collection
    Set xLoaders = New Collection
    xLoaders.Add xLoader
    xLoaders.Add xLoader
        
    xSut.RegisterUsingLoader xLoaders
    
    
    'Act:
    Assert.AreEqual CLng(2), xLoader.LoadToContainerCallsCount

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Load")
Private Sub LoadSingle()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xLoader As VbaDITestRegistrationLoader
    Set xLoader = New VbaDITestRegistrationLoader
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    xSut.RegisterUsingLoader xLoader
    
    
    'Act:
    Assert.AreEqual CLng(1), xLoader.LoadToContainerCallsCount

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub ReturnsFirstRegisteredObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(New VbaDITestObject)
    xSut.Register xReg
    
    Set xReg = VbaDI.Instance(New ReturnsZeroTestObject)
    xSut.Register xReg
    
    Set xReg = VbaDI.ForInterface("AnyName") _
        .Use(New VbaDITestObject)
    xSut.Register xReg
    
    Set xReg = VbaDI.ForInterface("AnyName") _
        .Use(New ReturnsZeroTestObject)
    xSut.Register xReg
    
    'Act:
    Dim xTestObj As VbaDITestObject
    Set xTestObj = xSut.Resolve("AnyName")

    'Assert:
    Assert.IsTrue xTestObj.ReturnInputValue(5) = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RetrievesValueDependency()
    On Error GoTo TestFail
    
    'Arrange:
    Const xValueID As String = "ConnectionString"
    Const xExpected As String = "Test Connection String"
    
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.AddValueDependency xValueID
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(xTO).DependsOnValue(xValueID, xExpected)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    xSut.Register xReg
    
    'Act:
    Set xTO = xSut.Resolve(TypeName(xTO))
    
    Dim xActual As String
    xActual = xTO.InjectedValueDependency(xValueID)
    
    'Assert:
    Assert.AreEqual xExpected, xActual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RetrievesFirstValueDependency()
    On Error GoTo TestFail
    
    'Arrange:
    Const xValueID As String = "ConnectionString"
    Const xExpected As String = "Test Connection String"
    
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.AddValueDependency xValueID
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(xTO) _
        .DependsOnValue(xValueID, xExpected) _
        .DependsOnValue(xValueID, "Not " & xExpected)
        
    xSut.Register xReg
    
    'Act:
    Set xTO = xSut.Resolve(TypeName(xTO))
    
    Dim xActual As String
    xActual = xTO.InjectedValueDependency(xValueID)
    
    'Assert:
    Assert.AreEqual xActual, xExpected
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RetrievesValueRegisteredWithObject()
    On Error GoTo TestFail
    
    'Arrange:
    
    Const xValueID As String = "ConnectionString"
    Const xExpected As String = "Test Connection String"
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.AddValueDependency xValueID
    
    xSut.Register VbaDI.Instance(xTO).DependsOnValue(xValueID, xExpected)

    'Act:
    Set xTO = xSut.Resolve(TypeName(xTO))
    
    Dim xActual As String
    xActual = xTO.InjectedValueDependency(xValueID)

    'Assert:
    Assert.AreEqual xExpected, xActual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RegisterOrderInstanceThenInterface()
    On Error GoTo TestFail

    'Arrange:

    Dim xSut As IVbaDIFluentRegistration
    Set xSut = VbaDI.Instance(New VbaDITestObject).ForInterface("AnyThing")

    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    xContainer.Register xSut
    
    'Act:
    Dim xTestObj As VbaDITestObject
    Set xTestObj = xContainer.Resolve("AnyThing")

    'Assert:
    Assert.IsTrue TypeOf xTestObj Is VbaDITestObject

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RegisterSameInstanceTwiceWithDifferentInterfaces()
    On Error GoTo TestFail

    'Arrange:

    Dim xTestObj As VbaDITestObject
    Set xTestObj = New VbaDITestObject

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register VbaDI.Instance(xTestObj) _
        .ForInterface("AnyThing")
    xSut.Register VbaDI.Instance(xTestObj) _
        .ForInterface("Nothing")
    
    'Act:
    Dim IAnything As VbaDITestObject
    Dim INothing As VbaDITestObject
    Set IAnything = xSut.Resolve("AnyThing")
    Set INothing = xSut.Resolve("Nothing")

    'Assert:
    IAnything.InstanceID = 5
    Assert.IsTrue INothing.InstanceID = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RegisterInstanceWith2InterfacesInline()
    On Error GoTo TestFail

    'Arrange:

    Dim xTestObj As VbaDITestObject
    Set xTestObj = New VbaDITestObject

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register VbaDI.Instance(xTestObj) _
        .ForInterface("AnyThing") _
        .ForInterface("Nothing")
    
    'Act:
    Dim IAnything As VbaDITestObject
    Dim INothing As VbaDITestObject
    Set IAnything = xSut.Resolve("AnyThing")
    Set INothing = xSut.Resolve("Nothing")

    'Assert:
    IAnything.InstanceID = 5
    Assert.IsTrue INothing.InstanceID = 5

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RegisteredAsInstanceAndThenAsInterfaceReturnsCorrectInstance()
    On Error GoTo TestFail
    
    'Arrange:
    Dim xTestObject1 As VbaDITestObject
    Set xTestObject1 = TS.CreateVbaDITestObject(1)
    
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.Instance(xTestObject1)

    Dim xTestObject2 As VbaDITestObject
    Set xTestObject2 = TS.CreateVbaDITestObject(2)
    
    Dim xReg2 As IVbaDIFluentRegistration
    Set xReg2 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(xTestObject2)

    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    xSut.Register xReg1
    xSut.Register xReg2

    'Act:
    Set xTestObject1 = xSut.Resolve(TypeName(New IVbaDITestInterface))
    
    Set xTestObject2 = xSut.Resolve(TypeName(New VbaDITestObject))
    

    'Assert:
    Assert.IsTrue xTestObject1.InstanceID = 1
    Assert.IsTrue xTestObject2.InstanceID = 1
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub RegisteredAsAliasWithImplementationAndThenAsInstance()
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim xTestObj1 As VbaDITestObject
    Set xTestObj1 = TS.CreateVbaDITestObject(1)
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(xTestObj1)

    Dim xTestObj2 As VbaDITestObject
    Set xTestObj2 = TS.CreateVbaDITestObject(2)
    Dim xReg2 As IVbaDIFluentRegistration
    Set xReg2 = VbaDI.Instance(xTestObj2)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()

    'Act:
    xSut.Register Array(xReg1, xReg2)
    
    Set xTestObj1 = xSut.Resolve(TypeName(New IVbaDITestInterface))
    
    Set xTestObj2 = xSut.Resolve(TypeName(New VbaDITestObject))
    
    'Assert:
    Assert.IsTrue xTestObj1.InstanceID = 1
    Assert.IsTrue xTestObj2.InstanceID = 1
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub Registered2AsAliasesAndAnInstanceOrder1()

    On Error GoTo TestFail
    
    Registered2AsAliasesWithImplementationAndThenAsInstanceImpl _
        ItemListToCollection(1, 2, 3), 1
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub Registered2AsAliasesAndAnInstanceOrder2()

    On Error GoTo TestFail
    
    Registered2AsAliasesWithImplementationAndThenAsInstanceImpl _
        ItemListToCollection(1, 3, 2), 1
        
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub Registered2AsAliasesAndAnInstanceOrder3()

    On Error GoTo TestFail
    
    Registered2AsAliasesWithImplementationAndThenAsInstanceImpl _
        ItemListToCollection(2, 1, 3), 2
        
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'This lengthy Test replicates a scenario that previously failed during integration testing
'@TestMethod("VbaDI.Integration")
Private Sub MultiLevelDependenciesWithProvidedObjectIDs()

    On Error GoTo TestFail
    
    'Arrange:
       
    Dim xTestObjMap As VbaDIMap
    Set xTestObjMap = VbaDIMap.Create()
    
    Dim xTestObjNum As Long
    
    For xTestObjNum = 1 To 4
        xTestObjMap.Add xTestObjNum, New VbaDITestObject
    Next
    
    Dim xObjRegData1 As TRegData
    LoadTRegData xObjRegData1, xTestObjMap.Item(1), 1, 1
    
    Dim xObject2RegData As TRegData
    LoadTRegData xObject2RegData, xTestObjMap.Item(2), 2, 2
        
    Dim xObject3RegData As TRegData
    LoadTRegData xObject3RegData, xTestObjMap.Item(3), 3, 3
        
    Dim xObject4RegData As TRegData
    LoadTRegData xObject4RegData, xTestObjMap.Item(4), 4, 4
    
    Dim xTestObj As VbaDITestObject
    
    Set xTestObj = xTestObjMap.Item(1)
    xTestObj.AddObjectIDDependency xObject2RegData.ObjectID
            
    Set xTestObj = xTestObjMap.Item(2)
    xTestObj.AddObjectIDDependency xObject3RegData.ObjectID
            
    Set xTestObj = xTestObjMap.Item(3)
    xTestObj.AddObjectIDDependency xObject4RegData.ObjectID
    xTestObj.AddValueDependency "TestValue" ', "12345"
            
    Dim xRegistrations As Collection
    Set xRegistrations = New Collection
    
    Dim xObject1Registration As IVbaDIFluentRegistration
    Set xObject1Registration = VbaDI.Instance(xObjRegData1.Instance, _
        xObjRegData1.ObjectID)
        
    
    Dim xObject2Registration As IVbaDIFluentRegistration
    Set xObject2Registration = VbaDI.Instance(xObject2RegData.Instance, _
        xObject2RegData.ObjectID)
    
    Dim xObject3Registration As IVbaDIFluentRegistration
    Set xObject3Registration = VbaDI.Instance(xObject3RegData.Instance, _
        xObject3RegData.ObjectID) _
        .DependsOnValue("TestValue", "12345")
    
    
    Dim xObject4Registration As IVbaDIFluentRegistration
    Set xObject4Registration = VbaDI.Instance(xObject4RegData.Instance, _
        xObject4RegData.ObjectID)
    
    xRegistrations.Add xObject1Registration
    xRegistrations.Add xObject2Registration
    xRegistrations.Add xObject3Registration
    xRegistrations.Add xObject4Registration
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    xSut.Register xRegistrations
    
    'Act:
    Dim xResolvedObject1 As VbaDITestObject
    Set xResolvedObject1 = xSut.Resolve(xObjRegData1.ObjectID)
    
    Dim xResolvedObject2 As VbaDITestObject
    Set xResolvedObject2 = xSut.Resolve(xObject2RegData.ObjectID)
    
    Dim xResolvedObject3 As VbaDITestObject
    Set xResolvedObject3 = xSut.Resolve(xObject3RegData.ObjectID)

    'Assert:
    Assert.IsTrue TypeOf xResolvedObject1 Is VbaDITestObject
    Assert.IsTrue TypeOf xResolvedObject2 Is VbaDITestObject
    Assert.IsTrue TypeOf xResolvedObject3 Is VbaDITestObject
    
    Dim xObj As VbaDITestObject
    Set xObj = xResolvedObject1.InjectedObject(xObject2RegData.ObjectID)
    Assert.IsTrue Not xObj Is Nothing
    Assert.AreEqual "2", xObj.InstanceID
    
    Set xObj = xResolvedObject2.InjectedObject(xObject3RegData.ObjectID)
    Assert.IsTrue Not xObj Is Nothing
    Assert.AreEqual "3", xObj.InstanceID
    Assert.AreEqual "12345", xObj.InjectedValueDependency("TestValue")

    Set xObj = xResolvedObject3.InjectedObject(xObject4RegData.ObjectID)
    Assert.IsTrue Not xObj Is Nothing
    Assert.AreEqual "4", xObj.InstanceID

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub LoadTRegData(ByRef pT As TRegData, ByVal pObj As Object, _
    Optional ByVal pObjID As String = vbNullString, _
    Optional ByVal pInstID As String = vbNullString)
    
    Set pT.Instance = pObj
    pT.ObjectID = TypeName(pObj)
    If pObjID <> vbNullString Then
        pT.ObjectID = pObjID
        pT.Instance.ObjectID = pObjID
    End If
    pT.Instance.InstanceID = pInstID
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub NonClassNameObjectIDs()
    On Error GoTo TestFail

    'Arrange:
    Const xObjID As String = "Nonsense"
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.ObjectID = xObjID
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(xTO, xTO.ObjectID)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg
    
    'Act:
    Dim xObj As VbaDITestObject
    Set xObj = xSut.Resolve(xTO.ObjectID)

    'Assert:
    Assert.AreEqual TypeName(xTO), TypeName(xObj)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub SimpleCaseCallsComposeSingleton()
    On Error GoTo TestFail
    
    'Even if there are no dependencies to inject,
    'the container is obligated to call 'Compose' (if possible) in
    'order to execute any initialization code

    'Arrange:
    Const xObjID As String = "Nonsense"
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.ObjectID = xObjID
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(xTO, xTO.ObjectID)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg
    
    'Act:
    Dim xObj As VbaDITestObject
    Set xObj = xSut.Resolve(xTO.ObjectID)

    'Assert:
    Assert.AreEqual CLng(1), xObj.InjectDependenciesCalls

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Transient")
Private Sub SimpleCaseCallsComposeTransient()
    On Error GoTo TestFail
    
    'Even if there are no dependencies to inject,
    'the container is obligated to call 'Compose' (if possible) in
    'order to execute any initialization code

    'Arrange:
    Const xObjID As String = "Nonsense"
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.ObjectID = xObjID
    
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.Instance(xTO, xTO.ObjectID).AsTransient()
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    xSut.Register xReg
    
    'Act:
    Dim xObj As VbaDITestObject
    Set xObj = xSut.Resolve(xTO.ObjectID)

    'Assert:
    Assert.AreEqual CLng(1), xObj.InjectDependenciesCalls

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub SameClassDifferentObjectIDs()
    On Error GoTo TestFail

    'Arrange:
    Const xObjID As String = "Nonsense"
    Const xObjID2 As String = "Crazy"
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    Dim xReg As IVbaDIFluentRegistration
    
    Dim xTO As VbaDITestObject
    Set xTO = New VbaDITestObject
    xTO.ObjectID = xObjID
    Set xReg = VbaDI.Instance(xTO, xTO.ObjectID)
    xTO.InstanceID = 1
    
    xSut.Register xReg
    
    Dim xTO2 As VbaDITestObject
    Set xTO2 = New VbaDITestObject
    xTO2.ObjectID = UCase$(xObjID2)
    Set xReg = VbaDI.Instance(xTO2, xObjID2)
    xTO2.InstanceID = 2
   
    xSut.Register xReg
    
    'Act:
    Dim xObj As VbaDITestObject
    Set xObj = xSut.Resolve(xTO2.ObjectID)

    'Assert:
    Dim xActual As Long
    xActual = xObj.InstanceID
    Assert.AreEqual CLng(2), xActual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Registered2AsAliasesWithImplementationAndThenAsInstanceImpl( _
    ByVal pItemOrder As Collection, ByVal pExpectedInstanceNumber As Long)
    
    'Arrange:
    Dim xIVbaDITestInterface As VbaDITestObject
    Dim xIVbaDITestInterface2 As VbaDITestObject
    Dim xTestObject As VbaDITestObject
    
    Dim xItems As Collection
    Set xItems = CreateTestObjectRegistrationsTrio(xIVbaDITestInterface, xTestObject, _
        xIVbaDITestInterface2)
    
    Dim xSut As IVbaDIContainer
    Set xSut = RegisterInOrder(xItems, pItemOrder)
    
    'Act:
    Set xIVbaDITestInterface = _
        xSut.Resolve(TypeName(New IVbaDITestInterface))
        
    Set xIVbaDITestInterface2 = _
        xSut.Resolve(TypeName(New IVbaDITestInterface2))
    
    'Assert:
    Assert.AreEqual CStr(pExpectedInstanceNumber), xIVbaDITestInterface.InstanceID
    Assert.IsTrue xTestObject.InstanceID = "2"
    Assert.IsTrue CStr(pExpectedInstanceNumber), xIVbaDITestInterface2.InstanceID
    
End Sub
'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub OneCallToRegistrationIDs()
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim xTestObj1 As VbaDITestObject
    Set xTestObj1 = TS.CreateVbaDITestObject(1)
    xTestObj1.AddValueDependency "TestVal"
    Dim xReg1 As IVbaDIFluentRegistration
    Set xReg1 = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(xTestObj1) _
        .DependsOnValue("TestVal", "ValueString")
        

    Dim xTestObj2 As VbaDITestObject
    Set xTestObj2 = TS.CreateVbaDITestObject(2)
    xTestObj2.AddObjectIDDependency TypeName(New IVbaDITestInterface)
    Dim xReg2 As IVbaDIFluentRegistration
    Set xReg2 = VbaDI.Instance(xTestObj2, "SecondTestObject")

    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    'Act:
    xSut.Register xReg1
    xSut.Register xReg2
    
    Set xTestObj2 = xSut.Resolve("SecondTestObject")
    Set xTestObj1 = xTestObj2.InjectedObject(TypeName(New IVbaDITestInterface))
    
    
    'Assert:
    Assert.AreEqual CLng(1), xTestObj1.GetDependencyIDCalls, "xTestObj1 failed"
    Assert.AreEqual CLng(1), xTestObj2.GetDependencyIDCalls, "xTestObj2 failed"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'Strange test...was the result of a failing use case during integration testing
'@TestMethod("VbaDI.Container.Resolve.Singleton")
Private Sub ValueDependencySurvivesAfterNextRegistration()
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
        
    xTestObj1.AddValueDependency "TestVal"

    Dim xTestObj2 As VbaDITestObject
    Set xTestObj2 = TS.CreateVbaDITestObject(2)
    Dim xReg2 As IVbaDIFluentRegistration
    Set xReg2 = VbaDI.Instance(xTestObj2, "SecondTestObject")
    
    xTestObj2.AddObjectIDDependency TypeName(New IVbaDITestInterface)
    
    Dim xSut As IVbaDIContainer
    Set xSut = VbaDI.CreateContainer()
    
    'Dim xVDs As Variant
    'Set xVDs = util.GetElement(, VbaDIKey.ValueDependencies)

    'Act:
    xSut.Register xReg1, xReg2
    
    Set xTestObj2 = xSut.Resolve("SecondTestObject")
    Dim xFirst As VbaDITestObject
    Set xFirst = xTestObj2.InjectedObject(TypeName(New IVbaDITestInterface))
    
    'Assert:
    Dim xActual As String
    xActual = xFirst.InjectedValueDependency("TestVal")
    Assert.AreEqual xExpected, xActual
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.MergeRegistration")
Private Sub MergeRegistrationCopiesValueDependencies()
    On Error GoTo TestFail

    'Arrange:
    Dim xReg As VbaDIRegistration
    
    Dim xRecord1 As VbaDIMap
    Set xRecord1 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject")
    Set xRecord1 = TS.AddInterfaceID(xRecord1, "ISomething")
    Set xReg = RegistrationCode.CreateRegistration(xRecord1)
    
    Dim xReg2 As VbaDIRegistration
    
    Dim xRecord2 As VbaDIMap
    Set xRecord2 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject2")
    Set xRecord2 = TS.AddInterfaceID(xRecord2, "ISomething")
    Set xReg2 = RegistrationCode.CreateRegistration(xRecord2)
    
    Dim xExpected As String
    xExpected = "FirstVal"
    Set xReg2 = TS.AddValueDependency(xReg2, "FirstKey", xExpected)
    

    'Act:
    
    Dim xResult As VbaDIRegistration
    Set xResult = RegistrationCode.MergeRegistration(xReg, xReg2)

    'Assert:
    Dim xActual As Variant
    Dim xKey As String
    xKey = "FirstKey"
    If xResult.ValueDependencies.Exists(xKey) Then
        xActual = xResult.ValueDependencies.Item(xKey)
        Assert.AreEqual xExpected, CStr(xActual)
    Else
        Assert.Fail "Value dependency does not exist"
    End If
    
    Assert.AreEqual xExpected, CStr(xActual)
    

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.MergeRegistration")
Private Sub MergeRegistrationMergesValueDependencies()
    On Error GoTo TestFail

    'Arrange:
    Dim xReg As VbaDIRegistration
    
    Dim xRecord1 As VbaDIMap
    Set xRecord1 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject")
    Set xRecord1 = TS.AddInterfaceID(xRecord1, "ISomething")
    Set xReg = RegistrationCode.CreateRegistration(xRecord1)
    
    Dim xExpected As String
    xExpected = "FirstVal"
    
    Set xReg = TS.AddValueDependency(xReg, "FirstKey", xExpected)
    
    Dim xReg2 As VbaDIRegistration
    
    Dim xRecord2 As VbaDIMap
    Set xRecord2 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject2")
    Set xRecord2 = TS.AddInterfaceID(xRecord2, "ISomething")
    Set xReg2 = RegistrationCode.CreateRegistration(xRecord2)
    Set xReg2 = TS.AddValueDependency(xReg2, "SecondKey", "SecondVal")
    
    'Act:
    Dim xResult As VbaDIRegistration
    Set xResult = RegistrationCode.MergeRegistration(xReg, xReg2)

    'Assert:
    Dim xActual As Variant
    
    If xResult.ValueDependencies.Exists("FirstKey") Then
        xActual = xResult.ValueDependencies.Item("FirstKey")
        Assert.AreEqual xExpected, CStr(xActual)
    Else
        Assert.Fail "Depdendency Key not found"
    End If
    
    Dim xSecondActual As String
    If xReg2.ValueDependencies.Exists("SecondKey") Then
        xSecondActual = xResult.ValueDependencies.Item("SecondKey")
        Assert.AreEqual xExpected, CStr(xActual)
    Else
        Assert.Fail "Depdendency Key not found"
    End If
    
    Assert.AreEqual "SecondVal", CStr(xSecondActual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDI.MergeRegistration")
Private Sub MergeRegistrationNoRepeatedDependencyIDs()
    On Error GoTo TestFail

    'Arrange:
    Dim xReg As VbaDIRegistration
    
    Dim xRecord1 As VbaDIMap
    Set xRecord1 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject")
    Set xRecord1 = TS.AddInterfaceID(xRecord1, "ISomething")
    Set xReg = RegistrationCode.CreateRegistration(xRecord1)
    
    Set xReg = TS.AddValueDependency(xReg, "FirstKey", "FirstVal")
    
    Dim xReg2 As VbaDIRegistration
    Dim xRecord2 As VbaDIMap
    Set xRecord2 = RegistrationCode.CreateRegistrationRecord(New VbaDITestObject, "TestObject2")
    Set xRecord2 = TS.AddInterfaceID(xRecord2, "ISomething")
    Set xReg2 = RegistrationCode.CreateRegistration(xRecord2)
    
    Set xReg = TS.AddValueDependency(xReg, "SecondKey", "SecondVal")
    
    'Act:
    Dim xResult As VbaDIRegistration
    Set xResult = RegistrationCode.MergeRegistration(xReg, xReg2)

    'Assert:
    Dim xActual As Variant
    xActual = xResult.ValueDependencies.Count
    Assert.AreEqual CInt(2), CInt(xActual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


Private Function RegisterInOrder(ByVal pRegistrations As Collection, _
    ByVal pTestElements As Collection) As IVbaDIContainer
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()

    Dim xItem As Variant
    For Each xItem In pTestElements
        xContainer.Register pRegistrations.Item(xItem)
    Next
    Set RegisterInOrder = xContainer
End Function

Private Function CreateTestObjectRegistrationsTrio(ByRef pTestObject1 As VbaDITestObject, _
    ByRef pTestObject2 As VbaDITestObject, _
    ByRef pTestObject3 As VbaDITestObject) As Collection
    
    
    Dim xRegistrations As Collection
    Set xRegistrations = New Collection
    
    Set pTestObject1 = TS.CreateVbaDITestObject(1)
    Dim xReg As IVbaDIFluentRegistration
    Set xReg = VbaDI.ForInterface(TypeName(New IVbaDITestInterface)) _
        .Use(pTestObject1)
        
    xRegistrations.Add xReg
        
    Set pTestObject2 = TS.CreateVbaDITestObject(2)
    Set xReg = VbaDI.Instance(pTestObject2)
        
    xRegistrations.Add xReg

    Set pTestObject3 = TS.CreateVbaDITestObject(3)
    Set xReg = VbaDI.ForInterface(TypeName(New IVbaDITestInterface2)) _
        .Use(pTestObject3)
        
    xRegistrations.Add xReg
    
    Set CreateTestObjectRegistrationsTrio = xRegistrations
End Function

Private Function ItemListToCollection( _
    ParamArray pItems() As Variant) As Collection
    
    Dim results As Collection
    Set results = New Collection
    Dim vKey As Variant
    For Each vKey In pItems
        results.Add vKey
    Next
    Set ItemListToCollection = results
End Function


'@TestMethod("VbaDI.QueryCompose")
Private Sub UnexpectedErrorDuringRegistrationIDsTest()
    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_GETTING_REGISTRATIONIDS
    On Error GoTo TestFail

    'Arrange:
    VbaDIError.EnablePrintToImmediateWindow False
    
    Dim xTestObj As VbaDITestObject
    Set xTestObj = New VbaDITestObject
    xTestObj.CauseExceptionForRegistrationIDs True
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    xContainer.Register VbaDI.Instance(xTestObj)
    
    'Act:
    '@Ignore VariableNotUsed
    Dim xObj As Object
    '@Ignore AssignmentNotUsed
    Set xObj = xContainer.Resolve(TypeName(xTestObj))
    
    'Assert:
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Debug.Print "UnexpectedErrorDuringRegistrationIDsTest Error: " & Err.Description
        Resume Assert
    End If
End Sub

'@TestMethod("VbaDI.QueryCompose")
Private Sub UnexpectedErrorDuringComposeTest()
    Dim ExpectedError As Long
    ExpectedError = VbaDIError.ERROR_DURING_COMPOSE
    On Error GoTo TestFail

    'Arrange:
    VbaDIError.EnablePrintToImmediateWindow False
    
    Dim xTestObj As VbaDITestObject
    Set xTestObj = New VbaDITestObject
    xTestObj.CauseExceptionDuringCompose True
    
    Dim xContainer As IVbaDIContainer
    Set xContainer = VbaDI.CreateContainer()
    xContainer.Register VbaDI.Instance(xTestObj) _
        .DependsOnValue("TestVal", "1234")
    
    'Act:
    '@Ignore VariableNotUsed
    Dim xObj As Object
    '@Ignore AssignmentNotUsed
    Set xObj = xContainer.Resolve(TypeName(xTestObj))
    
    'Assert:
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Debug.Print "UnexpectedErrorDuringComposeTest Error: " & Err.Description
        Resume Assert
    End If
End Sub


