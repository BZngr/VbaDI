VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "An Inversion of Control (IoC) container for VBA"
'@ModuleDescription "An Inversion of Control (IoC) container for VBA"
'@Folder("VbaDI")

Option Explicit

Implements IVbaDIContainer

Private Type TVBAContainer
    Data As VbaDIData
    Resolver As ObjectResolver
End Type

Private this As TVBAContainer

Private Sub Class_Initialize()
    Set this.Data = DataProvider.GetAddInData(True)
    Set this.Resolver = New ObjectResolver
    Set this.Resolver.Copier = RegistrationCode
End Sub

'Note: used only by tests.  Ignored to avoid RD inspection flag when used as add-in
'@Ignore ProcedureNotUsed
Public Sub SetResolver(ByVal pResolver As ObjectResolver)
    Set this.Resolver = pResolver
End Sub

Private Sub IVbaDIContainer_Register(ParamArray pFluentRegistrations() As Variant)
    RegisterImpl util.Concat(New Collection, _
        ParamArrayCode.RemoveNesting(pFluentRegistrations))
End Sub

Private Sub RegisterImpl(ByVal pIRegistrations As Collection)
       
    If Not util.Every(pIRegistrations, _
        VbaDIFunc.Create(VbaDIFunc.IsIVbaDIFluentRegistration)) Then
        
        Err.Raise 5, TypeName(Me), _
            "Invalid Collection element: must be IVbaDIFluentRegistration instance"
    End If
    
    Dim xNewReg As VbaDIRegistration
    
    For Each xNewReg In util.Map(pIRegistrations, _
        VbaDIFunc.Create(VbaDIFunc.CreateRegistration))
        
        Set this.Data.RegistrationsByObjectID = _
            VbaDICode.InsertRegistration(this.Data.RegistrationsByObjectID, xNewReg)
            
        Set this.Data.ObjectIDsByInterfaceID = _
            VbaDICode.InsertInterfaceIDsForObjectID(this.Data.ObjectIDsByInterfaceID, xNewReg)
    Next
End Sub

Private Sub IVbaDIContainer_RegisterUsingLoader(ParamArray pLoaders() As Variant)
    LoadImpl util.Concat(New Collection, ParamArrayCode.RemoveNesting(pLoaders))
End Sub

Private Sub LoadImpl(ByVal pLoaders As Collection)
    
    If Not util.Every(pLoaders, VbaDIFunc.Create(VbaDIFunc.IsVbaDILoader)) Then
        
        Err.Raise 5, TypeName(Me), _
            "Invalid Collection element: must be IVbaDIRegistrationsLoader instance"
    End If
    
    Dim xLoader As IVbaDIRegistrationLoader
    For Each xLoader In pLoaders
        xLoader.LoadToContainer Me
    Next
End Sub

Private Function IVbaDIContainer_Resolve( _
    ByVal pRegistrationID As String) As Object
    
On Error GoTo ResolutionErrorExit

    Set IVbaDIContainer_Resolve = _
        this.Resolver.Resolve(this.Data, pRegistrationID)
    
On Error GoTo 0
    Exit Function

ResolutionErrorExit:
    If Not VbaDIError.IsVbaDIError(Err.Number) Then
        VbaDIError.RaiseError VbaDIError.ERROR_RESOLVING_OBJECT, _
            TypeName(Me), pRegistrationID
    End If
    
    VbaDIError.Rethrow Err
        
End Function


