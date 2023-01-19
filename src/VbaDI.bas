Attribute VB_Name = "VbaDI"
Attribute VB_Description = "Module exposing AddIn entry point methods"
'@ModuleDescription "Module exposing AddIn entry point methods"
'@Folder("VbaDI")
Option Explicit

'@EntryPoint
Public Function CreateContainer() As IVbaDIContainer
    
    Dim xContainer As VbaDIContainer
    Set xContainer = New VbaDIContainer
    Set CreateContainer = xContainer
    
End Function

'@EntryPoint
'@Description "Entry point to a fluent registration API to create a VbaDI registration data object"
Public Function Instance(ByVal pInstance As Object, _
    Optional ByVal pRegistrationID As String = vbNullString) As IVbaDIFluentRegistration
Attribute Instance.VB_Description = "Entry point to a fluent registration API to create a VbaDI registration data object"
    
    Dim xBuilder As IVbaDIFluentRegistration
    Set xBuilder = New VbaDIRegistrationRecordBuilder
    
    Dim xID As String
    xID = TypeName(pInstance)
    xID = IIf(pRegistrationID <> vbNullString, pRegistrationID, xID)
    
    Set Instance = xBuilder.Use(pInstance, xID)
End Function

'@EntryPoint
'@Description "Entry point to a fluent registration API for creating a VbaDI registration object"
Public Function ForInterface(ParamArray pRegistrationElements() As Variant) As IVbaDIFluentRegistration
Attribute ForInterface.VB_Description = "Entry point to a fluent registration API for creating a VbaDI registration object"
        
    Dim xBuilder As IVbaDIFluentRegistration
    Set xBuilder = New VbaDIRegistrationRecordBuilder
    
    Set ForInterface = xBuilder.ForInterface(pRegistrationElements)
End Function
