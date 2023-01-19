VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIDependencyProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Object used to inject/deliver dependencies as ID/Object or ID/Value pairs"
'@ModuleDescription "Object used to inject/deliver dependencies as ID/Object or ID/Value pairs"
'@Folder("VbaDI")

Option Explicit

Implements IVbaDIDependencyProvider

Private Type TVbaDIDependencyProvider
    ValueDependencies As VbaDIMap
    ObjectDependencies As VbaDIMap
End Type

Private this As TVbaDIDependencyProvider

Public Property Get Self() As VbaDIDependencyProvider
    Set Self = Me
End Property

Public Sub SetDependencies(ByVal pValueDependencies As VbaDIMap, _
    ByVal pObjectDependencies As VbaDIMap)
    
    Set this.ValueDependencies = util.Assign(VbaDIMap.Create(), pValueDependencies)
    Set this.ObjectDependencies = util.Assign(VbaDIMap.Create(), pObjectDependencies)
End Sub

Public Function ObjectFor(ByVal pDependencyID As Variant) As Object
    If Not this.ObjectDependencies.Exists(pDependencyID) Then
        VbaDIError.RaiseError VbaDIError.ERROR_DEPENDENCY_ID_NOT_REQUESTED, _
            TypeName(Me), pDependencyID
    End If
    
    Set ObjectFor = this.ObjectDependencies.Item(pDependencyID)
End Function

Public Function ValueFor(ByVal pDependencyID As Variant) As Variant
    If Not this.ValueDependencies.Exists(pDependencyID) Then
        VbaDIError.RaiseError VbaDIError.ERROR_DEPENDENCY_ID_NOT_REQUESTED, _
            TypeName(Me), pDependencyID
    End If
    
    ValueFor = this.ValueDependencies.Item(pDependencyID)
End Function

Private Function IVbaDIDependencyProvider_ValueFor( _
    ByVal pDependencyID As Variant) As Variant
    
    IVbaDIDependencyProvider_ValueFor = ValueFor(pDependencyID)
End Function

Private Function IVbaDIDependencyProvider_ObjectFor( _
    ByVal pDependencyID As Variant) As Object
    
    Set IVbaDIDependencyProvider_ObjectFor = ObjectFor(pDependencyID)
End Function


