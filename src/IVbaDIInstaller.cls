VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder("VbaDI")

Option Explicit

'@Description "Used to install registrations into an IVbaDIContainer instance"
Public Sub Install(ByVal pContainer As IVbaDIContainer)
Attribute Install.VB_Description = "Used to install registrations into an IVbaDIContainer instance"
End Sub
