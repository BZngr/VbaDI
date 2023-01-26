VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIRegistrationLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder("VbaDI")

Option Explicit

'@Description "Used to load registrations into an IVbaDIContainer instance"
Public Sub LoadToContainer(ByVal pContainer As IVbaDIContainer)
Attribute LoadToContainer.VB_Description = "Used to load registrations into an IVbaDIContainer instance"
End Sub
