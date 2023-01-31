VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistrationCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stateless object supporting VbaDIRegistration operations"
'@PredeclaredId
'@Folder("VbaDI")
'@ModuleDescription "Stateless object supporting VbaDIRegistration operations"
Option Explicit

Public Function CreateRegistrationRecord( _
    Optional ByVal pInstance As Object = Nothing, _
    Optional ByVal pObjectID As String = vbNullString) As Object
    
    If pInstance Is Nothing And pObjectID <> vbNullString Then
        Err.Raise 5, TypeName(Me), "Invalid ObjectId for 'Nothing' Instance"
    End If
    
    Set CreateRegistrationRecord = util.CreateDictionary()
    With CreateRegistrationRecord
        .Add VbaDIKey.Instance, pInstance
        .Add VbaDIKey.ObjectID, CreateRegistrationID(pInstance, pObjectID)
        .Add VbaDIKey.ValueDependencies, util.CreateDictionary()
        .Add VbaDIKey.InterfaceIDs, New Collection
        .Add VbaDIKey.IsSingleton, True
        .Add VbaDIKey.ObjectDependencyIDs, New Collection
    End With
End Function

Public Function CreateRegistrationID(ByVal pInstance As Object, ByVal pObjectID As String) As String
    CreateRegistrationID = pObjectID
    If pObjectID = vbNullString And Not pInstance Is Nothing Then
        CreateRegistrationID = TypeName(pInstance)
    End If
End Function

Public Function CreateRegistration(ByVal pObj As Object) As VbaDIRegistration

    Dim xData As Object
    If IsDictionary(pObj) Then
        Set xData = pObj
    ElseIf TypeOf pObj Is IVbaDIFluentRegistration Then
        Dim xBuilder As VbaDIRegistrationRecordBuilder
        Set xBuilder = pObj
        Set xData = xBuilder.Build()
    End If
    
    Set CreateRegistration = New VbaDIRegistration
    CreateRegistration.SetRecord xData '.Inner
End Function

Public Function IsDependencyFree( _
    ByVal pRegistration As VbaDIRegistration) As Boolean
    
    IsDependencyFree = Not TypeOf pRegistration.Instance Is IVbaDIQueryCompose _
        Or Not (util.Some(pRegistration.ObjectDependencyIDs) _
            Or util.Some(pRegistration.ValueDependencies))
End Function

Public Function GetDefaultClassFactory( _
    ByVal pRegistration As VbaDIRegistration) As IVbaDIDefaultFactory
    
    If Not TypeOf pRegistration.Instance Is IVbaDIDefaultFactory Then
        VbaDIError.RaiseError _
            VbaDIError.ERROR_DEFAULTFACTORY_NOT_IMPLEMENTED, _
            TypeName(Me), pRegistration.ObjectID
    End If

    Set GetDefaultClassFactory = pRegistration.Instance

End Function

'Note called from 'Object' (uses duck-typing) rather than explicitly this object
'@Ignore ProcedureNotUsed
Public Function CopyUsingNewInstance(ByVal pRegistration As VbaDIRegistration, _
    ByVal pNewInstance As Object) As VbaDIRegistration
    
    Set CopyUsingNewInstance = pRegistration
    
    Dim xNewRecord As Object
    Set xNewRecord = util.Assign(util.CreateDictionary(), pRegistration.AsRecord)
    
    Set xNewRecord = _
        util.SetElement(xNewRecord, VbaDIKey.Instance, pNewInstance)
    
    Set CopyUsingNewInstance = CreateRegistration(xNewRecord)
    
End Function

Public Function InsertRegistration(ByVal pRegistrationsByObjectID As Object, _
    ByVal pNewReg As VbaDIRegistration) As Object
    
    Set InsertRegistration = pRegistrationsByObjectID
    
    Dim xRegistrationsByObjectID As Object
    Set xRegistrationsByObjectID = _
        util.Assign(util.CreateDictionary(), pRegistrationsByObjectID)
        
    If Not util.HasKey(xRegistrationsByObjectID, pNewReg.ObjectID) Then
        Set xRegistrationsByObjectID = _
            util.SetElement(xRegistrationsByObjectID, pNewReg.ObjectID, pNewReg)
            
        Set InsertRegistration = xRegistrationsByObjectID
        Exit Function
    End If

    Dim xExistingRegistration As VbaDIRegistration
    Set xExistingRegistration = _
        util.GetElement(xRegistrationsByObjectID, pNewReg.ObjectID)
    
    If Not (xExistingRegistration.IsSingleton And pNewReg.IsSingleton) Then
        VbaDIError.RaiseError VbaDIError.ERROR_LIFESTYLE_MISMATCH, _
            TypeName(Me), xExistingRegistration.ObjectID
            
    End If
    
    Dim xMergedReg As VbaDIRegistration
    Set xMergedReg = MergeRegistration( _
        xExistingRegistration, pNewReg)
        
    Set xRegistrationsByObjectID.Item(xMergedReg.ObjectID) = xMergedReg

    Set InsertRegistration = xRegistrationsByObjectID
End Function

Public Function GetRegistration(ByVal pData As Object, _
    ByVal pKey As String) As VbaDIRegistration
    
    Set GetRegistration = Nothing
    If util.HasKey(pData, pKey) Then
        Set GetRegistration = util.GetElement(pData, pKey)
    End If
End Function

Public Function MergeRegistration( _
    ByVal pExistingRegistration As VbaDIRegistration, _
    ByVal pRegistrationToMerge As VbaDIRegistration) As VbaDIRegistration
           
    Dim xMergedRecord As Object
    Set xMergedRecord = _
        util.Assign(util.CreateDictionary(), pExistingRegistration.AsRecord)
    
    Dim xInterfaceIDs As Collection
    Set xInterfaceIDs = util.Concat( _
        util.GetElement(xMergedRecord, VbaDIKey.InterfaceIDs), _
        pRegistrationToMerge.InterfaceIDs)
    
    Set xMergedRecord = _
        util.SetElement(xMergedRecord, VbaDIKey.InterfaceIDs, xInterfaceIDs)
    
    Dim xExistingValuesMap As Object
    Set xExistingValuesMap = _
        util.GetElement(xMergedRecord, VbaDIKey.ValueDependencies)
    
    Dim xMergeValuesMap As Object
    Set xMergeValuesMap = util.Assign( _
        xExistingValuesMap, _
        util.PickBy(pRegistrationToMerge.ValueDependencies, _
            VbaDIFunc.Fn(EVbaDIFunc.NotInKeysOf, xExistingValuesMap)))
    
    Set xMergedRecord = util.SetElement( _
        xMergedRecord, VbaDIKey.ValueDependencies, xMergeValuesMap)
    
    Set MergeRegistration = Me.CreateRegistration(xMergedRecord)
    
End Function

Public Function AddInterfaceIDs(ByVal pReg As Object, _
    ByVal pIDs As Collection) As Object

    If Not util.Every(pIDs, VbaDIFunc.Fn(EVbaDIFunc.IsRegistrationID)) Then
        VbaDIError.RaiseError _
            VbaDIError.ERROR_INVALID_REGISTRATION_ID, _
            TypeName(Me)
    End If
    
    Dim xIDs As Collection
    Set xIDs = _
        util.Concat(pReg.Item(VbaDIKey.InterfaceIDs), pIDs)
    
    Dim xRegCopy As Object
    Set xRegCopy = util.Assign(util.CreateDictionary(), pReg)
    
    Set AddInterfaceIDs = _
        util.SetElement(xRegCopy, VbaDIKey.InterfaceIDs, xIDs)
    
End Function

Public Function IsDictionary(ByVal pObj As Object) As Boolean
    IsDictionary = TypeName(pObj) = "Dictionary"
End Function


