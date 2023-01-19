VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamArrayTestSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Tests"
Option Explicit

Private mForwards As Long
Private mResult As Variant

Public Function CreateNestedParamArray(ByVal pNestingLevel As Long, ParamArray pArr()) As Variant
    If pNestingLevel = 0 Then
        CreateNestedParamArray = pArr
        Exit Function
    End If
    
    mForwards = 1
    Nest pNestingLevel, pArr
    
    CreateNestedParamArray = mResult
End Function

Private Sub Nest(ByVal pGoal As Long, ParamArray pArr() As Variant)
    
    If mForwards = pGoal Then
        mResult = pArr
        Exit Sub
    End If
    
    mForwards = mForwards + 1
    Nest pGoal, pArr
End Sub

