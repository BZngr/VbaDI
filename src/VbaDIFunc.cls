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

Implements IVbaDIFunc

Private Type TVbaDIFunc
    FuncName As String
    FuncContext As Variant
End Type

Private this As TVbaDIFunc

Public Function Create(ByVal pFuncName As String, Optional ByVal pFuncContext As Variant = Empty) As IVbaDIFunc
    With New VbaDIFunc
        .FuncName = pFuncName
        If Not IsEmpty(pFuncContext) Then
            If TypeOf pFuncContext Is Object  Then
                Set .FuncContext = pFuncContext
            Else
                .FuncContext = pFuncContext
            End If
        End If
        Set Create = .Self
    End With
End Function

Public Property Get Self() As Object
    Set Self = Me
End Property

'@Ignore WriteOnlyProperty
Public Property Let FuncName(ByVal RHS As String)
    this.FuncName = RHS
End Property

'@Ignore WriteOnlyProperty
Public Property Let FuncContext(ByVal RHS As Variant)
    this.FuncContext = RHS
End Property

'@Ignore WriteOnlyProperty
Public Property Set FuncContext(ByVal RHS As Variant)
    Set this.FuncContext = RHS
End Property

Public Property Get IsIVbaDIFluentRegistration() As String
    IsIVbaDIFluentRegistration = "IsIVbaDIFluentRegistration"
End Property

Public Property Get IsVbaDILoader() As String
    IsVbaDILoader = "IsIVbaDILoader"
End Property

Public Property Get IsRegistrationID() As String
    IsRegistrationID = "IsRegistrationID"
End Property

Public Property Get IsSingleton() As String
    IsSingleton = "IsSingleton"
End Property

Public Property Get IsTransient() As String
    IsTransient = "IsTransient"
End Property

Public Property Get NotInKeys() As String
    NotInKeys = "NotInKeys"
End Property

Public Property Get CreateRegistration() As String
    CreateRegistration = "CreateRegistration"
End Property

Private Function IVbaDIFunc_Predicate(ByVal pVal As Variant, _
    ByVal pIndex As Variant, ByVal pColl As Variant) As Boolean
    
    IVbaDIFunc_Predicate = False
    
    Select Case this.FuncName
        Case Me.IsIVbaDIFluentRegistration
            IVbaDIFunc_Predicate = TypeOf pVal Is IVbaDIFluentRegistration
        Case Me.IsVbaDILoader
            IVbaDIFunc_Predicate = TypeOf pVal Is IVbaDIRegistrationLoader
        Case Me.IsRegistrationID
            On Error GoTo ErrorExit
            IVbaDIFunc_Predicate = Len(CStr(pVal)) > 0
            On Error GoTo 0
        Case Me.NotInKeys
            IVbaDIFunc_Predicate = Not this.FuncContext.Exists(pVal)
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported FuncName: " & this.FuncName
    End Select
    
    Exit Function
ErrorExit:
End Function

Private Function IVbaDIFunc_MapFunc(ByVal pVal As Variant, _
    ByVal pIndex As Long, ByVal pColl As Collection) As Variant

    Dim xResult As Variant

    Select Case this.FuncName
        Case Me.CreateRegistration
            Set xResult = RegistrationCode.CreateRegistration(pVal)
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported FuncName: " & this.FuncName
      End Select
    
    Set IVbaDIFunc_MapFunc = xResult

End Function

Private Function IVbaDIFunc_ValueKeyPredicate(ByVal pVal As Variant, _
    ByVal pKey As Variant) As Variant

    IVbaDIFunc_ValueKeyPredicate = Empty
    
    Select Case this.FuncName
        Case Me.IsSingleton
            ValidateIsVbaDIRegistration pVal
            IVbaDIFunc_ValueKeyPredicate = pVal.IsSingleton
        Case Me.IsTransient
            ValidateIsVbaDIRegistration pVal
            IVbaDIFunc_ValueKeyPredicate = Not pVal.IsSingleton
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported FuncName: " & this.FuncName
      End Select

End Function

Private Sub ValidateIsVbaDIRegistration(ByVal pValue As Variant)
    If TypeOf pValue Is VbaDIRegistration Then
        Exit Sub
    End If
    
    Err.Raise 5, TypeName(Me), "Invalid parameter Type"
End Sub


