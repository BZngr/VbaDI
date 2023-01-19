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
    RegistrationsByObjectID As VbaDIMap
    ObjectIDsByInterfaceID As VbaDIMap
    ResolvedSingletonsByRegistrationID As VbaDIMap
End Type

Private this As TVbaDIContainerData

Public Property Get RegistrationsByObjectID() As VbaDIMap
    If this.RegistrationsByObjectID Is Nothing Then
        Set this.RegistrationsByObjectID = VbaDIMap.Create()
    End If
    Set RegistrationsByObjectID = this.RegistrationsByObjectID
End Property
Public Property Set RegistrationsByObjectID(ByVal RHS As VbaDIMap)
    Set this.RegistrationsByObjectID = RHS
End Property

Public Property Get ObjectIDsByInterfaceID() As VbaDIMap
    If this.ObjectIDsByInterfaceID Is Nothing Then
        Set this.ObjectIDsByInterfaceID = VbaDIMap.Create()
    End If
    Set ObjectIDsByInterfaceID = this.ObjectIDsByInterfaceID
End Property
Public Property Set ObjectIDsByInterfaceID(ByVal RHS As VbaDIMap)
    Set this.ObjectIDsByInterfaceID = RHS
End Property

Public Property Set ResolvedSingletonsByRegistrationID(ByVal RHS As VbaDIMap)
    Set this.ResolvedSingletonsByRegistrationID = RHS
End Property
Public Property Get ResolvedSingletonsByRegistrationID() As VbaDIMap
    If this.ResolvedSingletonsByRegistrationID Is Nothing Then
        Set this.ResolvedSingletonsByRegistrationID = VbaDIMap.Create()
    End If
    Set ResolvedSingletonsByRegistrationID = this.ResolvedSingletonsByRegistrationID
End Property

