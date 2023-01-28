VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDITestSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")
Option Explicit

Public Function CreateVbaDITestObject( _
    Optional ByVal pInstanceNumber As Long = -1) As VbaDITestObject
    
    Dim xSut As VbaDITestObject
    Set xSut = New VbaDITestObject
    xSut.InstanceID = IIf(pInstanceNumber > 0, pInstanceNumber, 0)
    Set CreateVbaDITestObject = xSut
End Function

Public Function AddValueDependency(ByVal pRegistration As VbaDIRegistration, _
    ByVal pValueID As String, _
    ByVal pValue As Variant) As VbaDIRegistration

    Dim xNewRecord As VbaDIMap
    Set xNewRecord = util.Assign(VbaDIMap.Create(), pRegistration.AsRecord)
    
    Dim xVD As VbaDIMap
    Set xVD = util.Assign(VbaDIMap.Create(), util.GetElement(xNewRecord, VbaDIKey.ValueDependencies))
    If Not xVD.Exists(pValueID) Then
        Set xVD = util.SetElement(xVD, pValueID, pValue)
    End If
    
    Set xNewRecord = util.SetElement(xNewRecord, VbaDIKey.ValueDependencies, xVD)
    Set AddValueDependency = RegistrationCode.CreateRegistration(xNewRecord)
End Function

Public Property Get Resolver() As ObjectResolver
    Set Resolver = New ObjectResolver
    Set Resolver.Copier = Me
End Property

'Note called from 'Object' rather than explicitly this object
'@Ignore ProcedureNotUsed
Public Function CopyUsingNewInstance( _
    ByVal pRegistration As VbaDIRegistration, _
    ByVal pNewInstance As Object) As VbaDIRegistration
    
    If Not TypeOf pRegistration.Instance Is VbaDITestObject Then
        Err.Raise 5, TypeName(Me), "UnExpected Type: " & TypeName(pRegistration.Instance)
        
    End If
    
    Dim xCopy As VbaDIRegistration
    Set xCopy = RegistrationCode.CopyUsingNewInstance(pRegistration, pNewInstance)
    
    Dim xTestCopy As VbaDITestObject
    Set xTestCopy = xCopy.Instance
    
    Dim xOG As VbaDITestObject
    Set xOG = pRegistration.Instance
    xOG.ObjectID = pRegistration.ObjectID
    
    With xTestCopy
        .InstanceID = CLng(xOG.InstanceID) + 1
        Dim xID As Variant
        For Each xID In xOG.ObjectIDDependencies
            .AddObjectIDDependency xID
        Next
        For Each xID In xOG.ValueDependencies
            .AddValueDependency xID
        Next
        .ObjectID = xOG.ObjectID
    End With
    
    Set CopyUsingNewInstance = xCopy
End Function

Public Function AddInterfaceID(ByVal pRecord As VbaDIMap, _
    ByVal pID As String) As VbaDIMap

    Set AddInterfaceID = pRecord
    
    Dim xRecordCopy As VbaDIMap
    Set xRecordCopy = util.Assign(VbaDIMap.Create(), pRecord)
    
    Set xRecordCopy.Item(VbaDIKey.InterfaceIDs) = _
        util.Concat(xRecordCopy.Item(VbaDIKey.InterfaceIDs), pID)
        
    Set AddInterfaceID = xRecordCopy
End Function

Public Function Last(ByVal pArr As Variant) As Variant
    
    Dim xIndex As Long
    xIndex = UBound(pArr, 1)
    
    If TypeOf pArr(xIndex) Is Object  Then
        Set Last = pArr(xIndex)
    Else
        Last = pArr(xIndex)
    End If
End Function


