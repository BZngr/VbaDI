VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIDependencyProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Provides read-only accessors to VbaDI registered Objects/Values"
'@Interface
'@Exposed
'@Folder("VbaDI")
'@ModuleDescription "Provides read-only accessors to VbaDI registered Objects/Values"

Option Explicit

Public Function ValueFor(ByVal pValueID As Variant) As Variant
End Function

Public Function ObjectFor(ByVal pObjectID As Variant) As Object
End Function
