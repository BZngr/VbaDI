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

Public Function SetElement(ByVal pData As Object, _
    ByVal pKey As String, ByVal pVal As Variant) As Object
    
    If Not TypeOf pData Is VbaDIMap Then
        Err.Raise 5, TypeName(Me), "Invalid Object Type"
    End If
        
    Set SetElement = pData
    
    Dim xCopy As VbaDIMap
    Set xCopy = Assign(VbaDIMap.Create(), pData)
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

Public Function GetElement(ByVal pData As VbaDIMap, _
    ByVal pKey As String) As Variant
    
    If HasKey(pData, pKey) Then
        If IsObject(pData.Item(pKey)) Then
            Set GetElement = pData.Item(pKey)
        Else
            GetElement = pData.Item(pKey)
        End If
    End If

End Function

Public Function HasKey(ByVal pData As Variant, ByVal pKey As String) As Boolean
    
    HasKey = False
    If TypeOf pData Is VbaDIMap Then
        HasKey = pData.Exists(pKey)
        Exit Function
    End If
    
    If TypeOf pData Is Collection Or IsArray(pData) Then
        Dim xItem As Variant
        For Each xItem In pData
            If UCase$(xItem) = UCase$(pKey) Then
                HasKey = True
                Exit Function
            End If
        Next
    End If
    
End Function

Public Function HasAny(ByVal pData As Variant) As Boolean
    
    HasAny = False
    
    If TypeOf pData Is VbaDIMap Or TypeOf pData Is Collection Then
        HasAny = pData.Count > 0
        Exit Function
    End If
    
    If IsArray(pData) Then
        HasAny = Not UBound(pData, 1) < LBound(pData, 1)
    End If

End Function

Public Function Assign(ByVal pTarget As VbaDIMap, _
    ParamArray pDictionaries() As Variant) As VbaDIMap

    Dim xDict As VbaDIMap
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

Public Function Every(ByVal pColl As Collection, ByVal pFunc As IVbaDIFunc) As Boolean

    Every = False
    Dim xIdx As Long

    For xIdx = 1 To pColl.Count
        If Not pFunc.Predicate(pColl.Item(xIdx), xIdx, pColl) Then
            Exit Function
        End If
    Next
    
    Every = True
End Function

Public Function Concat(ByVal pTarget As Collection, ParamArray pElements() As Variant) As Variant

    Dim xElements As Variant
    xElements = ParamArrayCode.RemoveNesting(pElements)

    Dim xResult As Collection
    Set xResult = New Collection
    
    Dim xElement As Variant
    
    For Each xElement In pTarget
        xResult.Add xElement
    Next
    
    For Each xElement In xElements
        If TypeOf xElement Is VbaDIMap Then
            Err.Raise 5, TypeName(Me), "Unable to concat a 'VbaDIMap'"
        End If
        
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

Public Function PickBy(ByVal pMap As VbaDIMap, ByVal pFunc As IVbaDIFunc) As VbaDIMap

    Set PickBy = VbaDIMap.Create()

    Dim xKey As Variant
    Dim xValue As Variant
    
    For Each xKey In pMap.Keys()
        
        ToVariant pMap.Item(xKey), xValue
        
        If pFunc.PredicateValueAndKey(xValue, xKey) Then
            PickBy.Add xKey, xValue
        End If
    Next
End Function

Public Function Map(ByVal pColl As Collection, ByVal pFunc As IVbaDIFunc) As Collection
    
    Set Map = New Collection
    
    Dim xIdx As Long
    For xIdx = 1 To pColl.Count
        Map.Add pFunc.MapFunc(pColl.Item(xIdx), xIdx, pColl)
    Next
End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub ToVariant(ByVal pSource As Variant, ByRef pResult As Variant)
    If TypeOf pSource Is Object  Then
        Set pResult = pSource
    Else
        pResult = pSource
    End If
End Sub

