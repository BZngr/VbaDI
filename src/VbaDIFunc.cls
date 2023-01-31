VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VbaDI"
'@PredeclaredId
Option Explicit

Implements IVDFunc

Public Enum EVbaDIFunc
    IsIVbaDIFluentRegistration = 2023
    IsVbaDILoader
    IsRegistrationID
    IsSingleton
    IsTransient
    NotInKeysOf
    CreateRegistration
End Enum

Private Type TVbaDIFunc
    FuncID As Long
    FuncContext As Variant
    FuncIDToName As Object
End Type

Private this As TVbaDIFunc

Public Function Fn(ByVal pFuncID As EVbaDIFunc, Optional ByVal pFuncContext As Variant = Empty) As IVDFunc
    With New VbaDIFunc
        .FuncID = pFuncID
        If Not IsEmpty(pFuncContext) Then
            If TypeOf pFuncContext Is Object  Then
                Set .FuncContext = pFuncContext
            Else
                .FuncContext = pFuncContext
            End If
        End If
        Set Fn = .Self
    End With
End Function

Public Property Get Self() As Object
    Set Self = Me
End Property

'@Ignore WriteOnlyProperty
Public Property Let FuncID(ByVal RHS As Long)
    this.FuncID = RHS
End Property

'@Ignore WriteOnlyProperty
Public Property Let FuncContext(ByVal RHS As Variant)
    this.FuncContext = RHS
End Property

'@Ignore WriteOnlyProperty
Public Property Set FuncContext(ByVal RHS As Variant)
    Set this.FuncContext = RHS
End Property

Private Function IVDFunc_Predicate(ByVal pVal As Variant, _
    ByVal pIndex As Variant, ByVal pContainer As Variant) As Boolean
    
    IVDFunc_Predicate = False
    
    Select Case this.FuncID
        Case IsIVbaDIFluentRegistration
            IVDFunc_Predicate = TypeOf pVal Is IVbaDIFluentRegistration
        Case IsVbaDILoader
            IVDFunc_Predicate = TypeOf pVal Is IVbaDIRegistrationLoader
        Case IsRegistrationID
            On Error GoTo ErrorExit
            IVDFunc_Predicate = Len(CStr(pVal)) > 0
            On Error GoTo 0
        Case NotInKeysOf 'Not in Keys of the FuncContext
            IVDFunc_Predicate = Not util.HasKey(this.FuncContext, pVal)
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported IVDFunc_Predicate: " & _
                GetFuncName(this.FuncID)
    End Select
    
    Exit Function
ErrorExit:
End Function

Private Function IVDFunc_MapFunc(ByVal pVal As Variant, _
    ByVal pIndex As Long, ByVal pContainer As Variant) As Variant

    Dim xResult As Variant

    Select Case this.FuncID
        Case CreateRegistration
            Set xResult = RegistrationCode.CreateRegistration(pVal)
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported IVDFunc_MapFunc: " & _
                GetFuncName(this.FuncID)
      End Select
    
    Set IVDFunc_MapFunc = xResult

End Function

Private Function IVDFunc_ValueKeyPredicate(ByVal pValue As Variant, _
    ByVal pKey As Variant) As Variant

    IVDFunc_ValueKeyPredicate = Empty
    
    Select Case this.FuncID
        Case IsSingleton
            If TypeOf pValue Is VbaDIRegistration Then
                IVDFunc_ValueKeyPredicate = pValue.IsSingleton
            End If
        Case IsTransient
            If TypeOf pValue Is VbaDIRegistration Then
                IVDFunc_ValueKeyPredicate = Not pValue.IsSingleton
            End If
        Case NotInKeysOf 'Not in Keys of the FuncContext
            IVDFunc_ValueKeyPredicate = Not util.HasKey(this.FuncContext, pValue)
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported IVDFunc_ValueKeyPredicate: " & _
                GetFuncName(this.FuncID)
      End Select

End Function

Private Function GetFuncName(ByVal pID As Long) As String
    
    InitializeFuncNamesMap
    
    GetFuncName = "Unknown"
    If this.FuncIDToName.Exists(pID) Then
        GetFuncName = this.FuncIDToName.Item(pID)
    End If
End Function

Private Sub InitializeFuncNamesMap()

    If Not this.FuncIDToName Is Nothing Then
        Exit Sub
    End If
    
    Set this.FuncIDToName = CreateObject("Scripting.Dictionary")
    With this.FuncIDToName
        .Add IsIVbaDIFluentRegistration, "IsIVbaDIFluentRegistration"
        .Add IsVbaDILoader, "IsVbaDILoader"
        .Add IsRegistrationID, "IsRegistrationID"
        .Add IsSingleton, "IsSingleton"
        .Add IsTransient, "IsTransient"
        .Add NotInKeysOf, "NotInKeysOf"
        .Add CreateRegistration, "CreateRegistration"
    End With
    
End Sub

