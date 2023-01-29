VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Wraps record of a registered application object"
'@ModuleDescription "Wraps record of a registered application object"
'@Folder "VbaDI"
Option Explicit

Private Type TVbaDIRegistration
    Record As Object
End Type

Private this As TVbaDIRegistration

Public Sub SetRecord(ByVal pRegistrationRecord As Object)
    Set this.Record = pRegistrationRecord
End Sub

Public Property Get AsRecord() As Object
    Set AsRecord = this.Record
End Property

Public Property Get Instance() As Object
    Set Instance = util.GetElement(this.Record, VbaDIKey.Instance)
End Property

Public Property Get ObjectID() As String
    ObjectID = util.GetElement(this.Record, VbaDIKey.ObjectID)
End Property

Public Property Get ValueDependencies() As Object
    Set ValueDependencies = _
        util.GetElement(this.Record, VbaDIKey.ValueDependencies)
End Property

Public Property Get InterfaceIDs() As Collection
    Set InterfaceIDs = util.GetElement(this.Record, VbaDIKey.InterfaceIDs)
End Property

Public Property Get ObjectDependencyIDs() As Collection
    Set ObjectDependencyIDs = _
        util.GetElement(this.Record, VbaDIKey.ObjectDependencyIDs)
End Property

Public Property Get IsSingleton() As Boolean
    IsSingleton = util.GetElement(this.Record, VbaDIKey.IsSingleton)
End Property


