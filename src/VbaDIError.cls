VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaDIError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Object supporting VbaDI custom errors and error messages"
'@ModuleDescription "Object supporting VbaDI custom errors and error messages"
'@IgnoreModule UnderscoreInPublicClassModuleMember
'@Folder("VbaDI")
'@PredeclaredId
Option Explicit

Private Enum EVbaDIError
    [_FirstErrorMinusOne] = vbObjectError + 12345
    ElementNotRegistered
    IVbaDIDefaultFactoryNotImplemented
    DependencyIDNotRequested
    LifestyleMismatch
    RegistrationIncomplete
    InvalidRegistrationID
    UnexpectedErrorDuringComposeID
    UnexpectedErrorDuringRegistrationIDs
    ResolveObjectError
    [_LastErrorPlusOne]
End Enum

Private Type TVbaDIErrors
    EnumToFormatString As Object
    PrintToImmediate As Boolean
End Type

Private this As TVbaDIErrors

Private Sub Class_Initialize()
    this.PrintToImmediate = True
    Set this.EnumToFormatString = CreateObject("Scripting.Dictionary")
    With this.EnumToFormatString
        .Add ElementNotRegistered, "'{0}' not registered"
        
        .Add IVbaDIDefaultFactoryNotImplemented, " '{0}' does not implement " & _
            TypeName(New IVbaDIDefaultFactory)
            
        .Add DependencyIDNotRequested, " Dependency ID '{0}' not requested"
        
        .Add LifestyleMismatch, " Lifestyle mismatch registered for '{0}'"
        
        .Add RegistrationIncomplete, " Registration Incomplete for '{0}'"
        
        .Add ResolveObjectError, " Unexpected Error while resolving '{0}'"
        
        .Add InvalidRegistrationID, " Invalid RegistrationID. 'CStr(<RegistrationID>)' must succeed"
        
        .Add UnexpectedErrorDuringComposeID, " Unexpected error during '{0}:IQueryCompose.Compose'"
        
        .Add UnexpectedErrorDuringRegistrationIDs, " Unexpected error during '{0}:IQueryCompose.RegistrationIDs'"
    End With
End Sub

'Note: Ignored to avoid Inspection result when used as an AddIn.
'Note: PrintToImmediate flag is set/reset during testing only
'@Ignore ProcedureNotUsed
Public Sub EnablePrintToImmediateWindow(ByVal pEnable As Boolean)
    this.PrintToImmediate = pEnable
End Sub

Public Function IsVbaDIError(ByVal pErrorID As Long) As Boolean
    IsVbaDIError = pErrorID > EVbaDIError.[_FirstErrorMinusOne] And pErrorID < EVbaDIError.[_LastErrorPlusOne]
End Function

Public Sub Rethrow(ByVal pErr As ErrObject)
    pErr.Raise pErr.Number, pErr.Source, pErr.Description
End Sub
Public Property Get ERROR_ELEMENT_NOT_REGISTERED() As Long
    ERROR_ELEMENT_NOT_REGISTERED = ElementNotRegistered
End Property

Public Property Get ERROR_DEFAULTFACTORY_NOT_IMPLEMENTED() As Long
    ERROR_DEFAULTFACTORY_NOT_IMPLEMENTED = _
        IVbaDIDefaultFactoryNotImplemented
End Property

Public Property Get ERROR_DEPENDENCY_ID_NOT_REQUESTED() As Long
    ERROR_DEPENDENCY_ID_NOT_REQUESTED = DependencyIDNotRequested
End Property

Public Property Get ERROR_LIFESTYLE_MISMATCH() As Long
    ERROR_LIFESTYLE_MISMATCH = LifestyleMismatch
End Property

Public Property Get ERROR_REGISTRATION_INCOMPLETE() As Long
    ERROR_REGISTRATION_INCOMPLETE = RegistrationIncomplete
End Property

Public Property Get ERROR_RESOLVING_OBJECT() As Long
    ERROR_RESOLVING_OBJECT = ResolveObjectError
End Property

Public Property Get ERROR_INVALID_REGISTRATION_ID() As Long
    ERROR_INVALID_REGISTRATION_ID = InvalidRegistrationID
End Property

Public Property Get ERROR_DURING_COMPOSE() As Long
    ERROR_DURING_COMPOSE = UnexpectedErrorDuringComposeID
End Property

'Note: Ignored to avoid Inspection result when used as an AddIn.
'@Ignore ProcedureNotUsed
Public Property Get ERROR_GETTING_REGISTRATIONIDS() As Long
    ERROR_GETTING_REGISTRATIONIDS = UnexpectedErrorDuringRegistrationIDs
End Property

Public Sub RaiseError(ByVal pErrNumber As Long, _
    ByVal pSource As String, ParamArray pFormatValues() As Variant)
    
    Dim xMsg As String
    xMsg = FormatString(GetDescription(pErrNumber), pFormatValues)
        
    If this.PrintToImmediate Then
        Debug.Print "Error: " & Err.Source & xMsg
    End If
    
    Err.Raise pErrNumber, pSource, xMsg
End Sub

Private Function GetDescription(ByVal pErrorID As Long) As String
    GetDescription = "No Description available"
    If this.EnumToFormatString.Exists(pErrorID) Then
        GetDescription = this.EnumToFormatString.Item(pErrorID)
    End If
End Function

'Original version by Dan Terkildsen
'https://stackoverflow.com/questions/13011747/composite-string-formatting-ie-using-0-1-and-2-in-a-string-to-format-i
'Modifications:
'1. Names to avoid RD Inspection results when referenced as an AddIn
'2. Flatten ParamArray allowing FormatString to be later in a forwarded call chain
'3. Fix bug: Failed when xIndexedValue had no formatting directive (':')
Private Function FormatString(ByVal formatText As String, _
    ParamArray pIndexes() As Variant) As String
        
    Dim xPA As Variant
    xPA = pIndexes
    
    Dim xIndexedValues As Variant
    xIndexedValues = ParamArrayCode.RemoveNesting(xPA)
    
    Dim xResult As String
    xResult = formatText
    
    Dim xIndexedValue As Variant
    Dim xIndex As Long
    xIndex = 0

    Dim xFormat As String
    Dim xStart As Long
    For Each xIndexedValue In xIndexedValues
        xStart = InStr(xResult, "{" & xIndex & "}")
        If xStart < 1 Then
            xStart = InStr(xResult, "{" & xIndex & ":")
        End If

        If xStart > 0 Then
            Dim xEnd As Long
            xEnd = InStr(xStart, xResult, "}")

            Dim xFormattedValue As String

            Dim xFormatPos As Long
            xFormatPos = InStr(xStart, xResult, ":")
            If xFormatPos < xEnd And xFormatPos <> 0 Then
                xFormat = Mid$(xResult, xFormatPos + 1, xEnd - xFormatPos - 1)
                xFormattedValue = Format$(xIndexedValue(xIndex), xFormat)
            Else
                xFormattedValue = xIndexedValue
            End If

            xResult = Left$(xResult, xStart - 1) & _
                      xFormattedValue & _
                      Mid$(xResult, xEnd + 1)

        End If
        xIndex = xIndex + 1
    Next

    FormatString = xResult

End Function


