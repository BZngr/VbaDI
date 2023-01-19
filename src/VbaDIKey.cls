VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VbaDI")
'@PredeclaredId
Option Explicit

Private Type TVbaDIKeys
    Instance As String
    ObjectID As String
    ValueDependencies As String
    InterfaceIDs As String
    IsSingleton As String
    ObjectDependencyIDs As String
End Type

Private this As TVbaDIKeys

Private Sub Class_Initialize()
    this.Instance = "Instance"
    this.ObjectID = "ObjectID"
    this.ValueDependencies = "ValueDependencies"
    this.InterfaceIDs = "InterfaceIDs"
    this.IsSingleton = "IsSingleton"
    this.ObjectDependencyIDs = "ObjectDependencyIDs"
End Sub

Public Property Get Instance() As String
    Instance = this.Instance
End Property

Public Property Get ObjectID() As String
    ObjectID = this.ObjectID
End Property

Public Property Get ValueDependencies() As String
    ValueDependencies = this.ValueDependencies
End Property

Public Property Get InterfaceIDs() As String
    InterfaceIDs = this.InterfaceIDs
End Property

Public Property Get IsSingleton() As String
    IsSingleton = this.IsSingleton
End Property

Public Property Get ObjectDependencyIDs() As String
    ObjectDependencyIDs = this.ObjectDependencyIDs
End Property

