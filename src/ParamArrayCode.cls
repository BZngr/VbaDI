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
'@Folder("ParamArrayCode")
Option Explicit

'Note:
'If the initial ParamArray contains a single 1-D array,
'ParamArrayCode.Flatten returns the contained array
'rather than an array containing a 1-D array

'@Description "Returns an Array free of extra 'layers' caused by forwarding a ParamArray as the sole ParamArray argument"
Public Function Flatten( _
    ParamArray pParamArray() As Variant) As Variant
Attribute Flatten.VB_Description = "Returns an Array free of extra 'layers' caused by forwarding a ParamArray as the sole ParamArray argument"
    
    Dim xArray As Variant
    xArray = pParamArray
    
    Do While Not IsInitialParamArray(xArray)
        xArray = xArray(LBound(xArray, 1))
    Loop
    
    Flatten = xArray
End Function

Private Function IsInitialParamArray(ByVal pArr As Variant) As Boolean
    
    IsInitialParamArray = False
        
    'A forwarded ParamArray results in a 1-D, single element array
    'Exception: See Note above
    
    If Dimensions(pArr) <> 1 Or Length(pArr) <> 1 Then
        'Anything but a 1-D single element array is the initial ParamArray
        IsInitialParamArray = True
    ElseIf Not IsArray(pArr(LBound(pArr, 1))) Then
        'An array starting with a value/object is the initial ParamArray content
        IsInitialParamArray = True
    End If
End Function

Private Function Dimensions(ByVal pArr As Variant) As Long

    'https://excelatfinance.com/xlf/xlf-big-array.php
    Const MAX_ARRAY_DIMENSIONS As Long = 60
    
    Dimensions = 0
    Dim xDim As Long
    For xDim = 1 To MAX_ARRAY_DIMENSIONS
        If Not HasDimension(pArr, xDim) Then
            Dimensions = xDim - 1
            Exit Function
        End If
    Next

End Function

Private Function HasDimension( _
    ByVal pArr As Variant, ByVal pDimNumber As Long) As Boolean

    HasDimension = False
    
On Error GoTo NoDimension
    HasDimension = UBound(pArr, pDimNumber) - LBound(pArr, pDimNumber) > -1
    
    Exit Function
    
NoDimension:
End Function

Private Function Length(ByVal pArr As Variant) As Long
    Length = UBound(pArr, 1) - LBound(pArr, 1) + 1
End Function
