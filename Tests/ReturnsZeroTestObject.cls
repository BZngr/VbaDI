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

Implements IVbaDIRegistrationLoader

Private mLoader As VbaDITestRegistrationLoader

Private Sub Class_Initialize()
    Set mLoader = New VbaDITestRegistrationLoader
End Sub

Private Sub IVbaDIRegistrationLoader_LoadToContainer(ByVal pContainer As IVbaDIContainer)
    Dim xLoader  As IVbaDIRegistrationLoader
    Set xLoader = mLoader
    xLoader.LoadToContainer pContainer
End Sub

Private Function IVbaDITestInterface_ReturnInputValue(ByVal inputVal As Long) As Long
    IVbaDITestInterface_ReturnInputValue = ReturnInputValue(inputVal)
End Function

'@Ignore ParameterNotUsed
Public Function ReturnInputValue(ByVal inputVal As Long) As Long
    ReturnInputValue = 0
End Function

