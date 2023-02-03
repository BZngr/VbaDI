VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIRegistrationRecordBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VbaDI"

Option Explicit

Implements IVbaDIFluentRegistration

Private Type TVbaDIRegistrationBuilder
    Registration As Object
End Type

Private this As TVbaDIRegistrationBuilder

Private Sub Class_Initialize()
    Set this.Registration = RegistrationCode.CreateRegistrationRecord()
End Sub

Private Function IVbaDIFluentRegistration_ForInterface( _
    ParamArray pInterfaceIDs() As Variant) As IVbaDIFluentRegistration
    
    Set IVbaDIFluentRegistration_ForInterface = Me
    
    Dim xInterfaces As Variant
    xInterfaces = ParamArrayCode.Flatten(pInterfaceIDs)
    
    Dim xElement As Variant
    For Each xElement In xInterfaces
        Set this.Registration = RegistrationCode.AddInterfaceIDs( _
            this.Registration, util.Concat(New Collection, xElement))
    Next
       
End Function

Private Function IVbaDIFluentRegistration_Use(ByVal pInstance As Object, _
    Optional ByVal pRegistrationID As String = vbNullString) As IVbaDIFluentRegistration
    
    If pInstance Is Nothing Then
        Err.Raise 5, TypeName(Me), "Object Parameter equals 'Nothing'"
    End If
    
    Set IVbaDIFluentRegistration_Use = Me
    
    'The first object instance registered/set 'wins',
    If Not util.GetElement(this.Registration, VbaDIKey.Instance) Is Nothing Then
        Exit Function
    End If
    
    Set this.Registration.Item(VbaDIKey.Instance) = pInstance
    this.Registration.Item(VbaDIKey.ObjectID) = _
        RegistrationCode.CreateRegistrationID(pInstance, pRegistrationID)
       
End Function

Private Function IVbaDIFluentRegistration_DependsOnValue( _
    ByVal pRegistrationID As String, _
    ByVal pValue As Variant) As IVbaDIFluentRegistration

    Set IVbaDIFluentRegistration_DependsOnValue = Me
    
    Dim xData As Object
    Set xData = util.GetElement(this.Registration, VbaDIKey.ValueDependencies)
    
    If Not util.HasKey(xData, pRegistrationID) Then
        Set this.Registration.Item(VbaDIKey.ValueDependencies) = _
            util.SetElement(xData, pRegistrationID, pValue)
    End If

End Function

Private Function IVbaDIFluentRegistration_AsSingleton() As IVbaDIFluentRegistration
    
    Set IVbaDIFluentRegistration_AsSingleton = Me
    
    Set this.Registration = _
        util.SetElement(this.Registration, VbaDIKey.IsSingleton, True)
End Function

Private Function IVbaDIFluentRegistration_AsTransient() As IVbaDIFluentRegistration
    
    Set IVbaDIFluentRegistration_AsTransient = Me
    
    Set this.Registration = _
        util.SetElement(this.Registration, VbaDIKey.IsSingleton, False)
End Function

Public Function Build() As Object

    Dim xInstance As Object
    Set xInstance = util.GetElement(this.Registration, VbaDIKey.Instance)

    If xInstance Is Nothing Then
        VbaDIError.RaiseError VbaDIError.ERROR_REGISTRATION_INCOMPLETE, _
            TypeName(Me), _
            "Concrete instance not set"
    End If
    
    If Not TypeOf xInstance Is IVbaDIQueryCompose Then
        'No dependency requirements published by the class
        Set Build = this.Registration
        Exit Function
    End If

    'IVbaDIQueryCompose is implemented - get the class' dependencyIDs
    
    Dim xQC As IVbaDIQueryCompose
    Set xQC = xInstance
    
On Error GoTo ErrorExit
    Dim xAllDependencyIDs As Collection
    Set xAllDependencyIDs = xQC.RegistrationIDs
On Error GoTo 0
    
    Dim xValueDependenciesMap As Object
    Set xValueDependenciesMap = _
        util.GetElement(this.Registration, VbaDIKey.ValueDependencies)
    
    Dim xClassDependencyIDs As Collection
    Set xClassDependencyIDs = util.Filter(xAllDependencyIDs, _
        VbaDIFunc.Fn(EVbaDIFunc.NotInKeysOf, xValueDependenciesMap))
       
    Set this.Registration = util.SetElement( _
        this.Registration, VbaDIKey.ObjectDependencyIDs, xClassDependencyIDs)

    Set Build = this.Registration
    
    Exit Function
    
ErrorExit:
    VbaDIError.RaiseError VbaDIError.ERROR_GETTING_REGISTRATIONIDS, _
        TypeName(Me), TypeName(Me)
End Function


