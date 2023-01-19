VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVbaDIDefaultFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Enables a Class Module to act as a Class Factory that invokes the default constructor"
'@ModuleDescription "Enables a Class Module to act as a Class Factory that invokes the default constructor"
'@Exposed
'@Interface
'@Folder("VbaDI")

Option Explicit

'@Description "Returns an Object instance created by the default constructor"
Public Function Create() As Object
Attribute Create.VB_Description = "Returns an Object instance created by the default constructor"
End Function
