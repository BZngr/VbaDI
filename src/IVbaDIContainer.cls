VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Public API of the VbaDI IoC container"
'@ModuleDescription "Public API of the VbaDI IoC container"
'@Folder("VbaDI")
'@Exposed
'@Interface

Option Explicit

'@Description "pIRegistration is a single IVbaDIRegistration or an Array/Collection of IVbaDIRegistrations"
'@Ignore ParameterNotUsed
Public Sub Register(ParamArray pRegistrations() As Variant)
Attribute Register.VB_Description = "pIRegistration is a single IVbaDIRegistration or an Array/Collection of IVbaDIRegistrations"
End Sub

'@Description "Composes a fully initialized Object by RegistrationID"
Public Function Resolve(ByVal pRegistrationID As String) As Object
Attribute Resolve.VB_Description = "Composes a fully initialized Object by RegistrationID"
End Function

'@Description "pLoaders is a single IVbaDIRegistrationLoader or an Array/Collection of IVbaDIRegistrationLoaders"
Public Sub RegisterUsingLoader(ParamArray pLoaders() As Variant)
Attribute RegisterUsingLoader.VB_Description = "pLoaders is a single IVbaDIRegistrationLoader or an Array/Collection of IVbaDIRegistrationLoaders"
End Sub
