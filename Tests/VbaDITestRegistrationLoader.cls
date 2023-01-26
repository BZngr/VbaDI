VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDITestRegistrationLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VbaDI.Tests")

Option Explicit

Implements IVbaDIRegistrationLoader

Private Type TVbaDITestRegistrationLoader
    LoadToContainerCallsCount As Long
End Type

Private this As TVbaDITestRegistrationLoader

Public Property Get LoadToContainerCallsCount() As Long
    LoadToContainerCallsCount = this.LoadToContainerCallsCount
End Property

Private Sub IVbaDIRegistrationLoader_LoadToContainer(ByVal pContainer As IVbaDIContainer)
    this.LoadToContainerCallsCount = this.LoadToContainerCallsCount + 1
End Sub
