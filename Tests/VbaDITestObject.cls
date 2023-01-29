VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDITestObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VbaDI.Tests")
Option Explicit

Implements IVbaDITestInterface

Implements IVbaDITestInterface2

Implements IVbaDIDefaultFactory

Implements IVbaDIQueryCompose

Private Type TVbaDITestObject
    InstanceID As String
    ObjectID As String
    Loader As VbaDITestRegistrationLoader
    ValueDependencies As Collection
    InjectedValueDependencies As Object
    ObjectIDDependencies As Collection
    InjectedObjectID2Object As Object
    GetDependencyIDsCount As Long
    InjectDependenciesCount As Long
    CauseErrorDuringCompose As Boolean
    CauseErrorOnRegistrationIDs As Boolean
End Type

Private this As TVbaDITestObject

Private Sub Class_Initialize()
    Set this.Loader = New VbaDITestRegistrationLoader
    this.ObjectID = TypeName(Me)
    Set this.ValueDependencies = New Collection
    Set this.ObjectIDDependencies = New Collection
    Set this.InjectedObjectID2Object = util.CreateDictionary()
    Set this.InjectedValueDependencies = util.CreateDictionary()
End Sub

Private Property Get IVbaDIQueryCompose_RegistrationIDs() As Collection
    
    Set IVbaDIQueryCompose_RegistrationIDs = New Collection
    
    If this.CauseErrorOnRegistrationIDs Then
        VbaDIError.RaiseError _
            VbaDIError.ERROR_GETTING_REGISTRATIONIDS, _
            TypeName(Me), TypeName(Me)
    End If
    
    this.GetDependencyIDsCount = this.GetDependencyIDsCount + 1
    Dim xID As Variant
    For Each xID In this.ValueDependencies
        IVbaDIQueryCompose_RegistrationIDs.Add xID
    Next
    For Each xID In this.ObjectIDDependencies
        IVbaDIQueryCompose_RegistrationIDs.Add xID
    Next
End Property

Private Sub IVbaDIQueryCompose_ComposeObject( _
    ByVal dependenciesProvider As IVbaDIDependencyProvider)
    
    If this.CauseErrorDuringCompose Then
        VbaDIError.RaiseError _
            VbaDIError.ERROR_DURING_COMPOSE, _
            TypeName(Me), TypeName(Me)
    End If
    
    this.InjectDependenciesCount = this.InjectDependenciesCount + 1
    Dim xID As Variant
    With dependenciesProvider
        For Each xID In this.ObjectIDDependencies
            this.InjectedObjectID2Object.Add xID, .ObjectFor(xID)
        Next
        For Each xID In this.ValueDependencies
            this.InjectedValueDependencies.Add xID, .ValueFor(xID)
        Next
    End With
    
End Sub

Private Function IVbaDITestInterface_ReturnInputValue(ByVal inputVal As Long) As Long
    IVbaDITestInterface_ReturnInputValue = ReturnInputValue(inputVal)
End Function

Private Function IVbaDIDefaultFactory_Create() As Object
    Set IVbaDIDefaultFactory_Create = New VbaDITestObject
End Function

Public Property Get GetDependencyIDCalls() As Long
    GetDependencyIDCalls = this.GetDependencyIDsCount
End Property

'@Ignore ProcedureNotUsed
Public Property Get InjectDependenciesCalls() As Long
    InjectDependenciesCalls = this.InjectDependenciesCount
End Property

Public Function ReturnInputValue(ByVal inputVal As Long) As Long
    ReturnInputValue = inputVal
End Function

Public Property Let ObjectID(ByVal RHS As String)
    this.ObjectID = RHS
End Property
Public Property Get ObjectID() As String
    ObjectID = this.ObjectID
End Property

Public Property Let InstanceID(ByVal RHS As String)
    this.InstanceID = RHS
End Property
Public Property Get InstanceID() As String
    InstanceID = this.InstanceID
End Property

Public Sub AddValueDependency(ByVal pID As String)
    this.ValueDependencies.Add pID
End Sub

Public Property Get ValueDependencies() As Collection
    Set ValueDependencies = this.ValueDependencies
End Property

Public Sub AddObjectIDDependency(ByVal pID As String)
    this.ObjectIDDependencies.Add pID
End Sub

Public Property Get ObjectIDDependencies() As Collection
    Set ObjectIDDependencies = this.ObjectIDDependencies
End Property

Public Function InjectedObject(ByVal pID As String) As Variant
    Set InjectedObject = this.InjectedObjectID2Object.Item(pID)
End Function

Public Function InjectedObjects() As Collection
    Dim xColl As Collection
    Set xColl = New Collection
    Dim xItem As Variant
    For Each xItem In this.InjectedObjectID2Object.Items()
        xColl.Add xItem
    Next
    Set InjectedObjects = xColl
End Function

Public Function InjectedValueDependency(ByVal pID As String) As Variant
    InjectedValueDependency = this.InjectedValueDependencies.Item(pID)
End Function

Public Sub CauseExceptionDuringCompose(ByVal pVal As Boolean)
    this.CauseErrorDuringCompose = pVal
End Sub

Public Sub CauseExceptionForRegistrationIDs(ByVal pVal As Boolean)
    this.CauseErrorOnRegistrationIDs = pVal
End Sub

'@Ignore EmptyMethod
Private Sub IVbaDITestInterface2_DoNothing()
End Sub


