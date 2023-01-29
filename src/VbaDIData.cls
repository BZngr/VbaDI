VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Hosts all IoC Container data"
'@ModuleDescription "Hosts all IoC Container data"
'@Folder("VbaDI")
Option Explicit

Private Type TVbaDIContainerData
    RegistrationsByObjectID As Object
    ObjectIDsByInterfaceID As Object
    ResolvedSingletonsByRegistrationID As Object
End Type

Private this As TVbaDIContainerData

Public Property Get RegistrationsByObjectID() As Object
    If this.RegistrationsByObjectID Is Nothing Then
        Set this.RegistrationsByObjectID = util.CreateDictionary()
    End If
    Set RegistrationsByObjectID = this.RegistrationsByObjectID
End Property
Public Property Set RegistrationsByObjectID(ByVal RHS As Object)
    Set this.RegistrationsByObjectID = RHS
End Property

Public Property Get ObjectIDsByInterfaceID() As Object
    If this.ObjectIDsByInterfaceID Is Nothing Then
        Set this.ObjectIDsByInterfaceID = util.CreateDictionary()
    End If
    Set ObjectIDsByInterfaceID = this.ObjectIDsByInterfaceID
End Property
Public Property Set ObjectIDsByInterfaceID(ByVal RHS As Object)
    Set this.ObjectIDsByInterfaceID = RHS
End Property

Public Property Get ResolvedSingletonsByRegistrationID() As Object
    If this.ResolvedSingletonsByRegistrationID Is Nothing Then
        Set this.ResolvedSingletonsByRegistrationID = util.CreateDictionary()
    End If
    Set ResolvedSingletonsByRegistrationID = this.ResolvedSingletonsByRegistrationID
End Property
Public Property Set ResolvedSingletonsByRegistrationID(ByVal RHS As Object)
    Set this.ResolvedSingletonsByRegistrationID = RHS
End Property

