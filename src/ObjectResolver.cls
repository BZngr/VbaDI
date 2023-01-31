VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VbaDI"
Option Explicit

Private Type TvdResolver
    Copier As Object
End Type

Private this As TvdResolver

Private Sub Class_Initialize()
    Set this.Copier = RegistrationCode
End Sub

'Supports Property Injection for Testing
'Note: Ignored to avoid Inspection result when used as an AddIn.
'@Ignore WriteOnlyProperty
'@Ignore ProcedureNotUsed
Public Property Set Copier(ByVal RHS As Object)
    If RHS Is Nothing Then
        Set this.Copier = RegistrationCode
    Else
        Set this.Copier = RHS
    End If
End Property

Public Function Resolve(ByVal pData As VbaDIData, _
    ByVal pRegistrationID As String) As Object
    
    Dim xReg As VbaDIRegistration
    Set xReg = VbaDICode.GetRegistration(pData, pRegistrationID)
    
    If RegistrationCode.IsDependencyFree(xReg) Then
        Set Resolve = ResolveSimpleCase(pData, xReg)
        Exit Function
    End If
    
    If xReg.IsSingleton Then
        Set Resolve = ResolveSingleton(pData, xReg, pRegistrationID)
    Else
        Set Resolve = ResolveTransient(pData, xReg)
    End If
End Function

Private Function ResolveSimpleCase(ByVal pData As VbaDIData, _
    ByVal pRegistration As VbaDIRegistration) As Object
    
    Dim xRegistrationToCompose As VbaDIRegistration
    Set xRegistrationToCompose = pRegistration
    
    If Not xRegistrationToCompose.IsSingleton Then
        Set xRegistrationToCompose = CreateTransientRegistration(pRegistration)
    End If
    
    Dim xObj As Object
    Set xObj = ComposeObject(xRegistrationToCompose)
    
    If pRegistration.IsSingleton Then
        Set pData.ResolvedSingletonsByRegistrationID = _
            VbaDICode.CacheResolvedInstance( _
                pData.ResolvedSingletonsByRegistrationID, pRegistration)
    End If
    
    Set ResolveSimpleCase = xObj
End Function

Private Function ResolveSingleton(ByVal pData As VbaDIData, _
    ByVal pRegistration As VbaDIRegistration, _
    ByVal pRegistrationID As String) As Object
      
    If util.HasKey(pData.ResolvedSingletonsByRegistrationID, pRegistrationID) Then
        Set ResolveSingleton = _
            util.GetElement( _
                pData.ResolvedSingletonsByRegistrationID, pRegistrationID)
        Exit Function
    End If
    
    Set ResolveSingleton = pRegistration.Instance
    
    Dim xObj As Object
    Set xObj = ResolveInstance(pData, pRegistration)
    
    Set pData.ResolvedSingletonsByRegistrationID = _
        VbaDICode.CacheResolvedInstance( _
            pData.ResolvedSingletonsByRegistrationID, pRegistration)
        
    Set ResolveSingleton = xObj
      
End Function

Private Function ResolveTransient(ByVal pData As VbaDIData, _
    ByVal pRegistration As VbaDIRegistration) As Object

    Dim xTransientRegistration As VbaDIRegistration
    Set xTransientRegistration = CreateTransientRegistration(pRegistration)

    'Even with no object dependencies, call 'Compose' to load
    'any value dependencies and execute initialization code
    If Not util.Some(xTransientRegistration.ObjectDependencyIDs) Then
        Set ResolveTransient = ComposeObject(xTransientRegistration)
        Exit Function
    End If
    
    Set ResolveTransient = ResolveInstance(pData, xTransientRegistration)

End Function

Private Function ResolveInstance(ByVal pData As VbaDIData, _
    ByVal pRegistration As VbaDIRegistration) As Object

    Set ResolveInstance = pRegistration.Instance
    
    'Even with no object dependencies, call 'Compose' to load
    'any value dependencies and execute initialization code
    If Not util.Some(pRegistration.ObjectDependencyIDs) Then
        Set ResolveInstance = ComposeObject(pRegistration)
        Exit Function
    End If
    
    Dim xRegistrationByDependencyID As Object
    Set xRegistrationByDependencyID = util.CreateDictionary()
    
    Dim xID As Variant
    For Each xID In pRegistration.ObjectDependencyIDs
        xRegistrationByDependencyID.Add xID, VbaDICode.GetRegistration(pData, xID)
    Next
    
    Dim xSingletonDependencies As Object
    Dim xTransientDependencies As Object
    LoadDependencyMaps pData, xRegistrationByDependencyID, _
        xSingletonDependencies, xTransientDependencies
        
    Dim xAllObjectDependencies As Object
    Set xAllObjectDependencies = util.Assign( _
        util.CreateDictionary(), _
        xSingletonDependencies, _
        xTransientDependencies)
    
    Set ResolveInstance = ComposeObject(pRegistration, xAllObjectDependencies)
End Function

Private Sub LoadDependencyMaps(ByVal pData As VbaDIData, _
    ByVal pRegistrationByDependencyID As Object, _
    ByRef pSingletonDependencies As Object, _
    ByRef pTransientDependencies As Object)
    
    Dim xSingletonRecords As Object
    Set xSingletonRecords = GetDependenciesMap(pRegistrationByDependencyID, IsSingleton)

    Set pSingletonDependencies = _
        ResolveSingletonObjectDependencyIDs(pData, xSingletonRecords)
    
    Dim xTransientRecords As Object
    Set xTransientRecords = GetDependenciesMap(pRegistrationByDependencyID, IsTransient)

    Set pTransientDependencies = _
        ResolveTransientObjectDependencyIDs(pData, xTransientRecords)
End Sub

Private Function GetDependenciesMap( _
    ByVal pRegistrationByDependencyIDs As Object, _
    ByVal pFuncID As EVbaDIFunc) As Object
    
    Dim xRecordsMap As Object
    Set xRecordsMap = util.PickBy(pRegistrationByDependencyIDs, _
        VbaDIFunc.Fn(pFuncID))

    Set GetDependenciesMap = xRecordsMap

End Function

Private Function ComposeObject(ByVal pRegistration As VbaDIRegistration, _
    Optional ByVal pObjectDependencies As Object = Nothing) As Object
    
    Set ComposeObject = pRegistration.Instance
    
    Dim xObjDependencies As Object
    If Not pObjectDependencies Is Nothing Then
        Set xObjDependencies = pObjectDependencies
    Else
        Set xObjDependencies = util.CreateDictionary()
    End If
    
    If Not TypeOf pRegistration.Instance Is IVbaDIQueryCompose Then
        Exit Function
    End If
    
    Dim xQC As IVbaDIQueryCompose
    Set xQC = pRegistration.Instance
    
    With New VbaDIDependencyProvider
        .SetDependencies pRegistration.ValueDependencies, xObjDependencies
On Error GoTo ComposeError
        xQC.ComposeObject .Self
On Error GoTo 0
    End With
    
    Exit Function
    
ComposeError:
    VbaDIError.RaiseError _
        VbaDIError.ERROR_DURING_COMPOSE, _
        TypeName(Me), pRegistration.ObjectID

End Function

Private Function ResolveSingletonObjectDependencyIDs( _
    ByVal pData As VbaDIData, _
    ByVal pRegistrationByObjectIDs As Object) As Object

    Dim xDepReg As VbaDIRegistration
    Dim xResolvedDependency As Object

    Dim xResolvedDependencies As Object
    Set xResolvedDependencies = util.CreateDictionary()

    Dim xID As Variant
    For Each xID In pRegistrationByObjectIDs.Keys()

        Set xDepReg = util.GetElement(pRegistrationByObjectIDs, xID)
        
        If util.HasKey(pData.ResolvedSingletonsByRegistrationID, xID) Then
            xResolvedDependencies.Add _
                xID, util.GetElement(pData.ResolvedSingletonsByRegistrationID, xID)
        Else
            Set xResolvedDependency = ResolveSingleton(pData, xDepReg, xID)

            Set pData.ResolvedSingletonsByRegistrationID = _
                VbaDICode.CacheResolvedInstance( _
                    pData.ResolvedSingletonsByRegistrationID, xDepReg)
            
            Set xResolvedDependencies = _
                util.SetElement(xResolvedDependencies, xID, xResolvedDependency)
        End If

    Next

    Set ResolveSingletonObjectDependencyIDs = xResolvedDependencies

End Function

Private Function ResolveTransientObjectDependencyIDs( _
    ByVal pData As VbaDIData, _
    ByVal pRegistrationByObjectIDs As Object) As Object
        
    Dim xDepReg As VbaDIRegistration
    Dim xResolvedDependency As Variant
    
    Dim xResolvedDependencies As Object
    Set xResolvedDependencies = util.CreateDictionary()
    
    Dim xID As Variant
    For Each xID In pRegistrationByObjectIDs.Keys()
        
        Set xDepReg = util.GetElement(pRegistrationByObjectIDs, xID)
        
        Set xResolvedDependency = ResolveTransient(pData, xDepReg)
        
        Set xResolvedDependencies = _
            util.SetElement(xResolvedDependencies, xID, xResolvedDependency)
    Next
    
    Set ResolveTransientObjectDependencyIDs = xResolvedDependencies

End Function

Private Function CreateTransientRegistration( _
    ByVal pRegistration As VbaDIRegistration) As VbaDIRegistration
    
    With RegistrationCode.GetDefaultClassFactory(pRegistration)
        Set CreateTransientRegistration = _
            this.Copier.CopyUsingNewInstance(pRegistration, .Create())
    End With

End Function


