Attribute VB_Name = "VbDIMapTests"
'@TestModule
'@Folder("Tests")

Option Explicit
Option Private Module


Private Assert As Rubberduck.AssertClass
'@Ignore VariableNotUsed
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
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

'@TestMethod("VbaDIMap")
Private Sub CaseSensitiveMap()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As VbaDIMap
    Set xSut = VbaDIMap.Create(0)
    
    xSut.Add "Name", "Tom"

    'Act:
    Dim xActual As Boolean
    xActual = xSut.Exists("NAME")

    'Assert:
    Assert.AreEqual False, xActual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDIMap")
Private Sub CaseInSensitiveMap()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As VbaDIMap
    Set xSut = VbaDIMap.Create()
    
    xSut.Add "Name", "Tom"

    'Act:
    Dim xActual As Boolean
    xActual = xSut.Exists("NAME")

    'Assert:
    Assert.AreEqual True, xActual

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDIMap")
Private Sub KeyPropertyObjectTest()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As VbaDIMap
    Set xSut = VbaDIMap.Create()
    
    xSut.Add "Name", New Collection

    'Act:
    xSut.Key("NAME") = "NewKey"

    'Assert:
    Assert.IsFalse xSut.Exists("NAME"), "Original key still exists"
    Assert.IsTrue xSut.Exists("NewKey"), "New key not found"
    Assert.IsTrue TypeOf xSut.Item("NewKey") Is Collection

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("VbaDIMap")
Private Sub KeyPropertyValueTest()
    On Error GoTo TestFail

    'Arrange:
    Dim xSut As VbaDIMap
    Set xSut = VbaDIMap.Create()
    
    xSut.Add "Name", "Tom"

    'Act:
    xSut.Key("NAME") = "NewKey"

    'Assert:
    Assert.IsFalse xSut.Exists("NAME"), "Original key still exists"
    Assert.IsTrue xSut.Exists("NewKey"), "New key not found"
    Assert.IsTrue xSut.Item("NewKey") = "Tom"

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("VbaDIMap")
Private Sub InvalidCompareModeLow()
    InvalidCompareModeTest -1
End Sub

'@TestMethod("VbaDIMap")
Private Sub InvalidCompareModeHigh()
    InvalidCompareModeTest 2
End Sub

Private Sub InvalidCompareModeTest(ByVal pMode As Long)
    Const ExpectedError As Long = 5
    On Error GoTo TestFail

    'Arrange:

    'Act:
    '@Ignore VariableNotUsed
    Dim xSut As VbaDIMap
    '@Ignore AssignmentNotUsed
    Set xSut = VbaDIMap.Create(pMode)

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

