VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIQueryCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Interface implemented by registered objects to identify/receive required RegistrationIDs"
'@ModuleDescription "Interface implemented by registered objects to identify/receive required RegistrationIDs"
'@Folder("VbaDI")
'@Exposed
'@Interface

Option Explicit

'@Description "Returns the dependency RegistrationIDs that an object requires
Public Property Get RegistrationIDs() As Collection
End Property

'@Description "Provides access to requested dependencies by RegistrationID"
Public Sub ComposeObject(ByVal pProvider As IVbaDIDependencyProvider)
Attribute ComposeObject.VB_Description = "Provides access to requested dependencies by RegistrationID"
End Sub
