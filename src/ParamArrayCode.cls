VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamArrayCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("VbaDI")
Option Explicit

'@Description "Removes nesting levels caused by forwarding a ParamArray as the ParamArray parameter one or more times"
Public Function RemoveNesting(ParamArray pParamArray() As Variant) As Variant
Attribute RemoveNesting.VB_Description = "Removes nesting levels caused by forwarding a ParamArray as the ParamArray parameter one or more times"
    
    'Calling this function adds one nesting level, so remove the
    'added level before evaluating the candidate
    RemoveNesting = RemoveNestingLevel(First(pParamArray))
    
End Function

Private Function RemoveNestingLevel(ByVal pArr As Variant) As Variant
    
    Dim xArrCandidate As Variant
    xArrCandidate = pArr
    
    If IsInitialParamArray(xArrCandidate) Then
        RemoveNestingLevel = xArrCandidate
        Exit Function
    End If
    
    RemoveNestingLevel = RemoveNesting(First(xArrCandidate))
        
End Function

Private Function IsInitialParamArray(ByVal pArr As Variant) As Boolean
    
    IsInitialParamArray = False
    
    'Empty ParamArray/Array
    If UBound(pArr, 1) - LBound(pArr, 1) = -1 Then
        IsInitialParamArray = True
        Exit Function
    End If
    
    'ParamArrays are always nested within single-dimension, single element arrays.
    'Finding any other kind of dimensioned array indicates that the
    'initial ParamArray content has been reached...and, it contained
    'a multi-dimensional array
    If CountOfDimensions(pArr) > 1 Then
        IsInitialParamArray = True
        Exit Function
    End If

    'Finding any value/object in the first position
    'indicates that the initial ParamArray content has been reached
    If Not IsArray(First(pArr)) Then
        IsInitialParamArray = True
    End If

End Function

Private Function CountOfDimensions(ByVal pArr As Variant) As Long

    'https://excelatfinance.com/xlf/xlf-big-array.php
    Const MAX_ARRAY_DIMENSIONS As Long = 60
    
    CountOfDimensions = 0
    Dim xDim As Long
    For xDim = 1 To MAX_ARRAY_DIMENSIONS
        If Not HasDimension(pArr, xDim) Then
            CountOfDimensions = xDim - 1
            Exit Function
        End If
    Next

End Function

Private Function HasDimension(ByVal pArr As Variant, ByVal pDimNumber As Long) As Boolean

    HasDimension = False
    
On Error GoTo NoDimension
    HasDimension = UBound(pArr, pDimNumber) - LBound(pArr, pDimNumber) > -1
    
    Exit Function
    
NoDimension:
End Function

Private Function First(ByVal pArr As Variant) As Variant
    
    Dim xIndex As Long
    xIndex = LBound(pArr, 1)
    
    If TypeOf pArr(xIndex) Is Object  Then
        Set First = pArr(xIndex)
    Else
        First = pArr(xIndex)
    End If
End Function

