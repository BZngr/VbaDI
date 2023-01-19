VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VbaDI"
'@PredeclaredId

Option Explicit

Private Type TDataProvider
    Data As VbaDIData
End Type

'@Ignore MoveFieldCloserToUsage
Private this As TDataProvider

Public Function GetAddInData(Optional ByVal pCreateNew As Boolean = False) As VbaDIData
    If this.Data Is Nothing Or pCreateNew Then
        Set this.Data = New VbaDIData
    End If
    
    Set GetAddInData = this.Data
End Function
