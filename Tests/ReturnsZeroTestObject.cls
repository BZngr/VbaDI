VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReturnsZeroTestObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VbaDI.Tests")
Option Explicit

Implements IVbaDITestInterface

Implements IVbaDIInstaller

Private mInstaller As VbaDITestInstaller

Private Sub Class_Initialize()
    Set mInstaller = New VbaDITestInstaller
End Sub

Private Sub IVbaDIInstaller_Install(ByVal pContainer As IVbaDIContainer)
    Dim xInstaller  As IVbaDIInstaller
    Set xInstaller = mInstaller
    xInstaller.Install pContainer
End Sub

Private Function IVbaDITestInterface_ReturnInputValue(ByVal inputVal As Long) As Long
    IVbaDITestInterface_ReturnInputValue = ReturnInputValue(inputVal)
End Function

'@Ignore ParameterNotUsed
Public Function ReturnInputValue(ByVal inputVal As Long) As Long
    ReturnInputValue = 0
End Function

