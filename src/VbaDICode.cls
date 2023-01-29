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

Public Function InsertRegistration(ByVal pRegistrationsByObjectID As Object, _
    ByVal pNewReg As VbaDIRegistration) As Object

    Set InsertRegistration = pRegistrationsByObjectID

    Set InsertRegistration = RegistrationCode.InsertRegistration( _
        pRegistrationsByObjectID, pNewReg)

End Function

Public Function InsertInterfaceIDsForObjectID( _
    ByVal pObjectIDsByInterfaceID As Object, _
    ByVal pRegistration As VbaDIRegistration) As Object
    
    Set InsertInterfaceIDsForObjectID = pObjectIDsByInterfaceID
    
    If pRegistration.InterfaceIDs.Count > 0 Then
        Dim xObj As Object
        Set xObj = util.Assign(util.CreateDictionary(), pObjectIDsByInterfaceID)
        
        Dim xItf As Variant
        For Each xItf In pRegistration.InterfaceIDs
            If Not util.HasKey(xObj, xItf) Then
                Set xObj = util.SetElement(xObj, xItf, pRegistration.ObjectID)
            End If
        Next
        Set InsertInterfaceIDsForObjectID = xObj
    End If
End Function

Public Function GetRegistration(ByVal pData As VbaDIData, _
    ByVal pRegistrationID As String) As VbaDIRegistration
    
    Set GetRegistration = Nothing
    
    With pData
        If util.HasKey(.RegistrationsByObjectID, pRegistrationID) Then
            Set GetRegistration = RegistrationCode.GetRegistration( _
                .RegistrationsByObjectID, pRegistrationID)
        
        ElseIf util.HasKey(.ObjectIDsByInterfaceID, pRegistrationID) Then
            Set GetRegistration = RegistrationCode.GetRegistration( _
                .RegistrationsByObjectID, _
                util.GetElement(.ObjectIDsByInterfaceID, pRegistrationID))
        End If
    End With
    
    If GetRegistration Is Nothing Then
        VbaDIError.RaiseError VbaDIError.ERROR_ELEMENT_NOT_REGISTERED, _
            TypeName(Me), pRegistrationID
    End If
End Function

Public Function CacheResolvedInstance( _
    ByVal pResolvedSingletonsByRegistrationID As Object, _
    ByVal pRegistration As VbaDIRegistration) As Object
    
    Set CacheResolvedInstance = pResolvedSingletonsByRegistrationID
    
    Dim xUpdatedCache As Object
    Set xUpdatedCache = _
        util.Assign(util.CreateDictionary(), pResolvedSingletonsByRegistrationID)
    
    Dim xID As Variant
    For Each xID In util.Concat( _
        pRegistration.InterfaceIDs, pRegistration.ObjectID)
        
        Set xUpdatedCache = _
            util.SetElement(xUpdatedCache, xID, pRegistration.Instance)
    Next
            
    Set CacheResolvedInstance = xUpdatedCache
End Function


