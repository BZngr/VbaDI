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

Private Type TPredicateFuncs
    FuncName As String
End Type

Private this As TPredicateFuncs

Public Function Create(ByVal pFuncName As String) As IVbaDIFunc
    With New VbaDIFunc
        .FuncName = pFuncName
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

Public Property Get CreateRegistration() As String
    CreateRegistration = "CreateRegistration"
End Property

Private Function IVbaDIFunc_Predicate(ByVal pVal As Variant, _
    ByVal pIndex As Long, ByVal pColl As Collection) As Boolean
    
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
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported FuncName: " & this.FuncName
    End Select
    
    Exit Function
ErrorExit:
End Function

Private Function IVbaDIFunc_PredicateValueAndKey( _
    ByVal pValue As Variant, ByVal pKey As String) As Boolean
    
    IVbaDIFunc_PredicateValueAndKey = False
    
    Select Case this.FuncName
        Case Me.IsSingleton
            ValidateIsVbaDIRegistration pValue
            IVbaDIFunc_PredicateValueAndKey = pValue.IsSingleton
        Case Me.IsTransient
            ValidateIsVbaDIRegistration pValue
            IVbaDIFunc_PredicateValueAndKey = Not pValue.IsSingleton
        Case Else
            Err.Raise 5, TypeName(Me), "Unsupported FuncName: " & this.FuncName
    End Select
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

Private Sub ValidateIsVbaDIRegistration(ByVal pValue As Variant)
    If TypeOf pValue Is VbaDIRegistration Then
        Exit Sub
    End If
    
    Err.Raise 5, TypeName(Me), "Invalid parameter Type"
End Sub


