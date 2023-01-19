VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDITestInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VbaDI.Tests")

Option Explicit

Implements IVbaDIInstaller

Private Type TVbaDITestInstaller
    InstallCallCount As Long
End Type

Private this As TVbaDITestInstaller

Public Property Get InstallCallsCount() As Long
    InstallCallsCount = this.InstallCallCount
End Property

Private Sub IVbaDIInstaller_Install(ByVal pContainer As IVbaDIContainer)
    this.InstallCallCount = this.InstallCallCount + 1
End Sub
