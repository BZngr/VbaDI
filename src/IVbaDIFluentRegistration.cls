VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIFluentRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Defines a fluent builder interface for the creation of VbaDI registration records"
'@ModuleDescription "Defines a fluent builder interface for the creation of VbaDI registration records"
'@Folder "VbaDI"
'@Exposed
'@Interface

Option Explicit

'@Description "Associates an instance with 1 or more interface implementations"
Public Function ForInterface(ParamArray pRegistrationElements() As Variant) As IVbaDIFluentRegistration
Attribute ForInterface.VB_Description = "Associates an instance with 1 or more interface implementations"
End Function

'@Description "Registers an instance"
Public Function Use(ByVal pInstance As Object, _
    Optional ByVal pRegistrationID As String = vbNullString) As IVbaDIFluentRegistration
Attribute Use.VB_Description = "Registers an instance"
End Function

'@Description "Registers a RegistrationID/Variant value pair"
Public Function DependsOnValue(ByVal pRegistrationID As String, _
    ByVal pValue As Variant) As IVbaDIFluentRegistration
Attribute DependsOnValue.VB_Description = "Registers a RegistrationID/Variant value pair"
End Function

'@Description "Container returns the same object instance for each dependency request"
Public Function AsSingleton() As IVbaDIFluentRegistration
Attribute AsSingleton.VB_Description = "Container returns the same object instance for each dependency request"
End Function

'@Description "Container returns a new object instance for each dependency request"
Public Function AsTransient() As IVbaDIFluentRegistration
Attribute AsTransient.VB_Description = "Container returns a new object instance for each dependency request"
End Function
