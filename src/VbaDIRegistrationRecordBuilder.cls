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
    Registration As VbaDIMap
End Type

Private this As TVbaDIRegistrationBuilder

Private Sub Initialize()
    If this.Registration Is Nothing Then
        Set this.Registration = RegistrationCode.CreateRegistrationRecord()
    End If
End Sub

Private Function IVbaDIFluentRegistration_ForInterface( _
    ParamArray pInterfaceIDs() As Variant) As IVbaDIFluentRegistration
    
    Set IVbaDIFluentRegistration_ForInterface = Me
    
    Initialize
    
    Dim xInterfaces As Variant
    xInterfaces = ParamArrayCode.RemoveNesting(pInterfaceIDs)
    
    Dim xElement As Variant
    For Each xElement In xInterfaces
        Set this.Registration = RegistrationCode.AddInterfaceIDs( _
            this.Registration, util.Concat(New Collection, xElement))
    Next
       
End Function

Private Function IVbaDIFluentRegistration_Use( _
    ByVal pInstance As Object, _
    Optional ByVal pRegistrationID As String = vbNullString) As IVbaDIFluentRegistration
    
    Set IVbaDIFluentRegistration_Use = Me
    
    Initialize
    
    If Not this.Registration.Item(VbaDIKey.Instance) Is Nothing Then
        Exit Function
    End If
    
    Dim xID As String
    If pRegistrationID <> vbNullString Then
        xID = pRegistrationID
    Else
        xID = TypeName(pInstance)
    End If
    
    Set this.Registration.Item(VbaDIKey.Instance) = pInstance
    this.Registration.Item(VbaDIKey.ObjectID) = xID
       
End Function

Private Function IVbaDIFluentRegistration_DependsOnValue( _
    ByVal pRegistrationID As String, _
    ByVal pValue As Variant) As IVbaDIFluentRegistration

    Dim xData As VbaDIMap
    Set xData = util.GetElement(this.Registration, VbaDIKey.ValueDependencies)
    
    If Not util.HasKey(xData, pRegistrationID) Then
        Set this.Registration.Item(VbaDIKey.ValueDependencies) = _
            util.SetElement(xData, pRegistrationID, pValue)
    End If

    Set IVbaDIFluentRegistration_DependsOnValue = Me

End Function

Private Function IVbaDIFluentRegistration_AsSingleton() As IVbaDIFluentRegistration
    Set this.Registration = util.SetElement(this.Registration, VbaDIKey.IsSingleton, True)
    Set IVbaDIFluentRegistration_AsSingleton = Me
End Function

Private Function IVbaDIFluentRegistration_AsTransient() As IVbaDIFluentRegistration
    Set this.Registration = util.SetElement(this.Registration, VbaDIKey.IsSingleton, False)
    Set IVbaDIFluentRegistration_AsTransient = Me
End Function

Public Function Build() As VbaDIMap

    Dim xInstance As Object
    Set xInstance = util.GetElement(this.Registration, VbaDIKey.Instance)

    If xInstance Is Nothing Then
        VbaDIError.RaiseError VbaDIError.ERROR_REGISTRATION_INCOMPLETE, _
            TypeName(Me), _
            "Concrete instance not set"
    End If

    If TypeOf xInstance Is IVbaDIQueryCompose Then

        Dim xValueDependencies As VbaDIMap
        Set xValueDependencies = util.GetElement( _
            this.Registration, VbaDIKey.ValueDependencies)
            
        Dim xObjectDependencyIDs As Collection
        Set xObjectDependencyIDs = _
            GetObjectDependencyIDs(xInstance, xValueDependencies)

        Set this.Registration = util.SetElement( _
            this.Registration, _
            VbaDIKey.ObjectDependencyIDs, _
            xObjectDependencyIDs)
    End If

    Set Build = this.Registration
End Function

Private Function GetObjectDependencyIDs(ByVal pQC As IVbaDIQueryCompose, ByVal pValDependencies As VbaDIMap) As Collection

    Set GetObjectDependencyIDs = New Collection
    
On Error GoTo ErrorExit
    Dim xDependencyIDs As Collection
    Set xDependencyIDs = pQC.RegistrationIDs
On Error GoTo 0

    Dim xDepID As Variant
    For Each xDepID In xDependencyIDs
        If Not util.HasKey(pValDependencies, xDepID) Then
            GetObjectDependencyIDs.Add CStr(xDepID)
        End If
    Next
    
    Exit Function
    
ErrorExit:
    VbaDIError.RaiseError VbaDIError.ERROR_GETTING_REGISTRATIONIDS, _
        TypeName(Me), TypeName(Me)


End Function


