VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Stateless utilities object"
'@PredeclaredId
'@ModuleDescription "Stateless utilities object"
'@Folder("VbaDI")

Option Explicit

Private Const TYPENAME_DICTIONARY As String = "Dictionary"

'For Case-sensitive use VbCompareMethod.vbBinaryCompare (0)
'Default is Case-insensitive comparisons (1)
Private Const DEFAULT_COMPARE_METHOD As Long = VbCompareMethod.vbTextCompare


Public Function SetElement(ByVal pData As Object, _
    ByVal pKey As String, ByVal pVal As Variant) As Object
    
    If Not IsDictionary(pData) Then
        Err.Raise 5, TypeName(Me), "Invalid Object Type"
    End If
        
    Set SetElement = pData
    
    Dim xCopy As Object
    Set xCopy = Assign(CreateDictionary(), pData)
    If Not HasKey(pData, pKey) Then
        xCopy.Add pKey, pVal
    Else
        If TypeOf pVal Is Object  Then
            Set xCopy.Item(pKey) = pVal
        Else
            xCopy.Item(pKey) = pVal
        End If
    End If
    
    Set SetElement = xCopy
End Function

Public Function GetElement(ByVal pData As Object, _
    ByVal pKey As String) As Variant
    
    Set GetElement = Nothing
    
    If HasKey(pData, pKey) Then
        If IsObject(pData.Item(pKey)) Then
            Set GetElement = pData.Item(pKey)
        Else
            GetElement = pData.Item(pKey)
        End If
    End If

End Function

Public Function HasKey(ByVal pData As Object, ByVal pKey As String) As Boolean
    
    HasKey = False
    If IsDictionary(pData) Then
        HasKey = pData.Exists(pKey)
    End If
    
End Function

Public Function Assign(ByVal pTarget As Object, _
    ParamArray pDictionaries() As Variant) As Object

    Dim xDict As Object
    Dim xKey As Variant
    Dim xItem As Variant
    For Each xItem In pDictionaries
        Set xDict = xItem
        For Each xKey In xDict.Keys()
            If pTarget.Exists(xKey) Then
                If IsObject(xDict.Item(xKey)) Then
                    Set pTarget.Item(xKey) = xDict.Item(xKey)
                Else
                    pTarget.Item(xKey) = xDict.Item(xKey)
                End If
            Else
                pTarget.Add xKey, xDict.Item(xKey)
            End If
        Next
    Next
    Set Assign = pTarget
End Function

Public Function First(ByVal pArr As Variant) As Variant
    
    Dim xIndex As Long
    xIndex = LBound(pArr, 1)
    
    If TypeOf pArr(xIndex) Is Object  Then
        Set First = pArr(xIndex)
    Else
        First = pArr(xIndex)
    End If
End Function

Public Function Some(ByVal pData As Variant, Optional ByVal pFunc As IVDFunc = Nothing) As Boolean
    
    Some = False
    Dim xIdx As Long
    
   If TypeOf pData Is Collection Then
        If pData.Count = 0 Then
            Exit Function
        End If
        
        If Not pFunc Is Nothing Then
            For xIdx = 1 To pData.Count
                If pFunc.Predicate(pData.Item(xIdx), xIdx, pData) Then
                    Some = True
                    Exit Function
                End If
            Next
        End If
        
        Some = True
        Exit Function
    End If
    
    If IsArray(pData) Then
        If IsEmptyArray(pData) Then
            Exit Function
        End If

        If Not pFunc Is Nothing Then
            For xIdx = LBound(pData, 1) To UBound(pData, 1)
                If pFunc.Predicate(pData(xIdx), xIdx, pData) Then
                    Some = True
                    Exit Function
                End If
            Next
        End If
    End If

End Function

Public Function Every(ByVal pData As Variant, Optional ByVal pFunc As IVDFunc = Nothing) As Boolean

    Every = False
    
    If TypeOf pData Is Collection Then
        If pData.Count = 0 Then
            Every = True
            Exit Function
        End If
        
        If Not pFunc Is Nothing Then
            Dim xIdx As Long
            For xIdx = 1 To pData.Count
                If Not pFunc.Predicate(pData.Item(xIdx), xIdx, pData) Then
                    Exit Function
                End If
            Next
            Every = True
        End If
        Exit Function
    End If
    
    If IsArray(pData) Then
        If IsEmptyArray(pData) Then
            Every = True
            Exit Function
        End If

        If Not pFunc Is Nothing Then
            For xIdx = LBound(pData, 1) To UBound(pData, 1)
                If Not pFunc.Predicate(pData(xIdx), xIdx, pData) Then
                    Exit Function
                End If
            Next
            Every = True
        End If
    End If
    
    Every = True
End Function

'Note: Flattens any Array/Collection found in pElements
Public Function Concat(ByVal pTarget As Collection, ParamArray pElements() As Variant) As Collection

    Dim xElements As Variant
    xElements = ParamArrayCode.RemoveNesting(pElements)

    Dim xResult As Collection
    Set xResult = New Collection
    
    Dim xElement As Variant
    
    For Each xElement In pTarget
        xResult.Add xElement
    Next
    
    For Each xElement In xElements
        If TypeOf xElement Is Collection Or IsArray(xElement) Then
            Dim xItem As Variant
            For Each xItem In xElement
                xResult.Add xItem
            Next
        Else
            xResult.Add xElement
        End If
    Next
    
    Set Concat = xResult
    
End Function

Public Function PickBy(ByVal pMap As Object, ByVal pFunc As IVDFunc) As Object

    Set PickBy = CreateDictionary()

    Dim xKey As Variant
    For Each xKey In pMap.Keys
        If pFunc.ValueKeyPredicate(pMap.Item(xKey), xKey) Then
            PickBy.Add CStr(xKey), pMap.Item(xKey)
        End If
    Next
End Function

Public Function Filter(ByVal pColl As Variant, ByVal pFunc As IVDFunc) As Collection

    Set Filter = New Collection
    Dim xIdx As Long
    If TypeOf pColl Is Collection Then
    
        For xIdx = 1 To pColl.Count
            If pFunc.Predicate(pColl.Item(xIdx), xIdx, pColl) Then
                Filter.Add pColl.Item(xIdx)
            End If
        Next
        Exit Function
    End If
    
    If IsArray(pColl) Then
        For xIdx = LBound(pColl, 1) To UBound(pColl, 1)
            If pFunc.Predicate(pColl(xIdx), xIdx, pColl) Then
                Filter.Add pColl(xIdx)
            End If
            Exit Function
        Next
    End If
    
    If IsDictionary(pColl) Then
        Dim xResult As Object
        Set xResult = CreateDictionary()
        For xIdx = 1 To pColl.Count
            Dim xT As Variant
            Dim xArr As Variant
            xArr = pColl.Keys
            xT = xArr(xIdx - 1)
            If pFunc.Predicate(pColl.Item(xT), xT, pColl) Then
                xResult.Add xT, pColl.Item(xT)
            End If
        Next
        Filter.Add xResult
        Exit Function
    End If
    
End Function

Public Function Map(ByVal pColl As Collection, ByVal pFunc As IVDFunc) As Collection
    
    Set Map = New Collection
    
    Dim xIdx As Long
    For xIdx = 1 To pColl.Count
        Map.Add pFunc.MapFunc(pColl.Item(xIdx), xIdx, pColl)
    Next
End Function

Private Function IsDictionary(ByVal pObj As Variant) As Boolean
    IsDictionary = False
    If TypeOf pObj Is Object  Then
        IsDictionary = TypeName(pObj) = TYPENAME_DICTIONARY
    End If
End Function

Private Function IsEmptyArray(ByVal pData As Variant) As Boolean
    IsEmptyArray = UBound(pData, 1) < LBound(pData, 1)
End Function

Public Function CreateDictionary(Optional ByVal pCompareMethod As Long = DEFAULT_COMPARE_METHOD) As Object
        
    'Excludes vbDatabaseCompare (2) as a valid value
    If Not (pCompareMethod = VbCompareMethod.vbBinaryCompare Or pCompareMethod = VbCompareMethod.vbTextCompare) Then
        Err.Raise 5, TypeName(Me), "Invalid Compare method value: " & (CStr(pCompareMethod))
    End If
    
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    
    CreateDictionary.CompareMode = IIf(pCompareMethod <> DEFAULT_COMPARE_METHOD, _
        pCompareMethod, DEFAULT_COMPARE_METHOD)
            
End Function


