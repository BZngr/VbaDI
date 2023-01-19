VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDICode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stateless object supporting VbaDI operations"
'@PredeclaredId
'@Folder("VbaDI")
'@ModuleDescription "Stateless object supporting VbaDI operations"
Option Explicit

Public Function InsertRegistration(ByVal pRegistrationsByObjectID As VbaDIMap, _
    ByVal pNewReg As VbaDIRegistration) As VbaDIMap

    Set InsertRegistration = pRegistrationsByObjectID

    Set InsertRegistration = RegistrationCode.InsertRegistration( _
        pRegistrationsByObjectID, pNewReg)

End Function

Public Function InsertInterfaceIDsForObjectID( _
    ByVal pObjectIDsByInterfaceID As VbaDIMap, _
    ByVal pRegistration As VbaDIRegistration) As VbaDIMap
    
    Set InsertInterfaceIDsForObjectID = pObjectIDsByInterfaceID
    
    If pRegistration.InterfaceIDs.Count > 0 Then
        Dim xCopy As VbaDIMap
        Set xCopy = util.Assign(VbaDIMap.Create(), pObjectIDsByInterfaceID)
        Dim xItf As Variant
        For Each xItf In pRegistration.InterfaceIDs
            If Not util.HasKey(xCopy, xItf) Then
                Set xCopy = _
                    util.SetElement(xCopy, xItf, pRegistration.ObjectID)
            End If
        Next
        Set InsertInterfaceIDsForObjectID = xCopy
    End If
End Function

Public Function GetRegistration(ByVal pData As VbaDIData, _
    ByVal pRegistrationID As String) As VbaDIRegistration
    
    With pData
        If util.HasKey(.RegistrationsByObjectID, pRegistrationID) Then
            Set GetRegistration = RegistrationCode.GetRegistration( _
                .RegistrationsByObjectID, pRegistrationID)
        
        ElseIf util.HasKey(.ObjectIDsByInterfaceID, pRegistrationID) Then
            Set GetRegistration = RegistrationCode.GetRegistration( _
                .RegistrationsByObjectID, _
                util.GetElement(.ObjectIDsByInterfaceID, pRegistrationID))
        Else
            VbaDIError.RaiseError VbaDIError.ERROR_ELEMENT_NOT_REGISTERED, _
                TypeName(Me), pRegistrationID
        End If
    End With
End Function


Public Function CacheResolvedInstance( _
    ByVal pResolvedSingletonsByRegistrationID As VbaDIMap, _
    ByVal pRegistration As VbaDIRegistration) As VbaDIMap
    
    Set CacheResolvedInstance = pResolvedSingletonsByRegistrationID
    
    Dim xUpdatedCache As VbaDIMap
    Set xUpdatedCache = _
        util.Assign(VbaDIMap.Create(), pResolvedSingletonsByRegistrationID)
    
    Dim xID As Variant
    For Each xID In util.Concat( _
        pRegistration.InterfaceIDs, pRegistration.ObjectID)
        
        Set xUpdatedCache = _
            util.SetElement(xUpdatedCache, xID, pRegistration.Instance)
    Next
            
    Set CacheResolvedInstance = xUpdatedCache
End Function

