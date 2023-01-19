VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "VbaDIMap wraps a Late-Bound Scripting.Dictionary"
'@Exposed
'@PredeclaredId
'@ModuleDescription "VbaDIMap wraps a Late-Bound Scripting.Dictionary"
'@IgnoreModule ProcedureNotUsed
'@Folder("VbaDI")

Option Explicit

'For Case-sensitive use VbCompareMethod.vbBinaryCompare (0)
'Default is Case-insensitive comparisons (1)
Private Const DEFAULT_COMPARE_METHOD As Long = VbCompareMethod.vbTextCompare

Private Type TVbaDIMap
    InnerDictionary As Object
End Type

Private this As TVbaDIMap

Public Function Create(Optional ByVal pCompareMethod As Long = DEFAULT_COMPARE_METHOD) As VbaDIMap
    
    'Excludes vbDatabaseCompare (2) as a valid value
    If Not (pCompareMethod = VbCompareMethod.vbBinaryCompare Or pCompareMethod = VbCompareMethod.vbTextCompare) Then
        Err.Raise 5, TypeName(Me), "Invalid Compare method value: " & (CStr(pCompareMethod))
    End If
    
    With New VbaDIMap
        .CompareMode = IIf(pCompareMethod <> DEFAULT_COMPARE_METHOD, _
            pCompareMethod, DEFAULT_COMPARE_METHOD)
            
        Set Create = .Self
    End With
End Function

Private Sub Class_Initialize()
    Set this.InnerDictionary = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get Self() As VbaDIMap
    Set Self = Me
End Property

Public Sub Add(ByVal pKey As Variant, ByVal pValue As Variant)
    this.InnerDictionary.Add pKey, pValue
End Sub

Public Function Exists(ByVal pKey As Variant) As Boolean
    Exists = this.InnerDictionary.Exists(pKey)
End Function

Public Function Items() As Variant
    Items = this.InnerDictionary.Items()
End Function

Public Function Keys() As Variant
    Keys = this.InnerDictionary.Keys()
End Function

Public Sub Remove(ByVal pKey As Variant)
    this.InnerDictionary.Remove pKey
End Sub

Public Sub RemoveAll()
    this.InnerDictionary.RemoveAll
End Sub

Public Property Get Count() As Long
    Count = this.InnerDictionary.Count
End Property

Public Property Get Item(ByVal pKey As Variant) As Variant

On Error GoTo UseSet:
    Item = this.InnerDictionary.Item(pKey)
On Error GoTo 0
    Exit Property

UseSet:
    Set Item = this.InnerDictionary.Item(pKey)
End Property
Public Property Let Item(ByVal pKey As Variant, ByVal RHS As Variant)
    this.InnerDictionary.Item(pKey) = RHS
End Property
Public Property Set Item(ByVal pKey As Variant, ByVal RHS As Variant)
    Set this.InnerDictionary.Item(pKey) = RHS
End Property

'@Ignore WriteOnlyProperty
Public Property Let Key(ByVal pKey As Variant, ByVal RHS As Variant)
    this.InnerDictionary.Key(pKey) = RHS
End Property

Public Property Get CompareMode() As Long
    CompareMode = this.InnerDictionary.CompareMode
End Property
Public Property Let CompareMode(ByVal RHS As Long)
    this.InnerDictionary.CompareMode = RHS
End Property


