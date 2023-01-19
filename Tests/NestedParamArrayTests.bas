Attribute VB_Name = "NestedParamArrayTests"
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

'@TestMethod("ParamArrayConverter")
Private Sub NonForwarded2ElementArray()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0, "A", "B")

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)
    

    'Assert:
    Assert.AreEqual "A", CStr(util.First(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub NonForwarded2ElementArray2Forwards()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0, "A", "B")

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)
    

    'Assert:
    Assert.AreEqual "A", CStr(util.First(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub NonForwarded1ElementArray()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport

    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0, "A")

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)
    

    'Assert:
    Assert.AreEqual "A", CStr(util.First(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub ExtractSingleLevel()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0, "A", "B")

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)
    

    'Assert:
    Assert.AreEqual "A", CStr(util.First(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub ExtractFromMultipleLevel()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(3, "A", "B", "C")

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)

    'Assert:
    Assert.AreEqual "C", CStr(TS.Last(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub Extract1DArrayFromMultipleLevel()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xArr As Variant
    xArr = Array("A", "B", "C")
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(3, xArr)

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)

    'Assert:
    Assert.AreEqual "C", CStr(TS.Last(xActual))

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub Extract2DArrayNonForwarded()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport
    
    Dim xArr() As Long
    ReDim xArr(0, 3)
    
    Dim xVal As Long
    xVal = 5
    Dim xD1 As Long
    xD1 = 0
    Dim xD2 As Long
    For xD2 = 0 To 3
        xVal = xD2 * 5
        xArr(xD1, xD2) = xVal
    Next
    
    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0, xArr)

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)

    'Assert:
    Assert.AreEqual CLng(10), xActual(0, 2)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub NonForwarded0ElementArray()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport

    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(0)

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)

    'Assert:
    Assert.AreEqual CLng(-1), UBound(xActual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ParamArrayConverter")
Private Sub Forwarded0ElementArray()
    On Error GoTo TestFail

    'Arrange:
    
    Dim xTO As ParamArrayTestSupport
    Set xTO = New ParamArrayTestSupport

    Dim xInput As Variant
    xInput = xTO.CreateNestedParamArray(3)

    'Act:
    Dim xActual As Variant
    xActual = ParamArrayCode.RemoveNesting(xInput)

    'Assert:
    Assert.AreEqual CLng(-1), UBound(xActual)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

