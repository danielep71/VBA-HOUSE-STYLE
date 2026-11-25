Attribute VB_Name = "M_Example"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Example
'------------------------------------------------------------------------------
' PURPOSE
'   Shows every layout rule of the house style in one small module
'
' WHY THIS EXISTS
'   Formatting is easier to copy from a worked example than from prose
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - a persisted greeting prefix
'     - greeting construction with validation
'
' ERROR POLICY
'   Public routines raise descriptive runtime errors unless explicitly
'   documented as best-effort
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   Replace EX_ with the project prefix of the target repository
'
' UPDATED
'   2026-01-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const EX_DEFAULT_PREFIX     As String = "Hello"     'Greeting used when none is set
    Private Const EX_MAX_NAME_LENGTH    As Long = 64            'Longest accepted name

    'Error numbers - one block per procedure family
    Private Const EX_ERR_BLANK_NAME     As Long = 1001          'Name is blank
    Private Const EX_ERR_LONG_NAME      As Long = 1002          'Name is too long

'------------------------------------------------------------------------------
' PUBLIC ENUMS
'------------------------------------------------------------------------------
    Public Enum EX_GreetingStyle
        EX_GreetingPlain = 0                                    'Prefix and name
        EX_GreetingFormal = 1                                   'Prefix, title-cased name
    End Enum

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mEX_Prefix                  As String               'Current greeting prefix, empty for default


'
'------------------------------------------------------------------------------
'
'                                  GREETINGS
'
'------------------------------------------------------------------------------
'

Public Function M_Greeting_Build( _
    ByVal PersonName As String, _
    Optional ByVal Style As EX_GreetingStyle = EX_GreetingPlain) As String
'
'------------------------------------------------------------------------------
'                                BUILD GREETING
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a greeting for one person
'
' WHY THIS EXISTS
'   Callers need one validated place that combines the prefix and the name
'
' INPUTS
'   PersonName
'     Name to greet. Leading and trailing spaces are ignored
'
'   Style
'     Optional greeting style
'       - EX_GreetingPlain   (default)
'       - EX_GreetingFormal
'
' RETURNS
'   The greeting text, for example "Hello Ada"
'
' BEHAVIOR
'   Normalizes the name, validates it, resolves the prefix, and joins them
'
' ERROR POLICY
'   Raises a descriptive runtime error if:
'     - PersonName is blank
'     - PersonName is longer than EX_MAX_NAME_LENGTH
'
' DEPENDENCIES
'   M_Greeting_GetPrefix
'
' NOTES
'   Typical call:
'
'       Debug.Print M_Greeting_Build("Ada", EX_GreetingFormal)
'
' UPDATED
'   2026-01-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Example.M_Greeting_Build"

    Dim NormalizedName          As String       'Trimmed person name
    Dim Prefix                  As String       'Resolved greeting prefix
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Ignore surrounding spaces
        NormalizedName = Trim$(PersonName)
    'Title-case the name for the formal style
        If Style = EX_GreetingFormal Then
            NormalizedName = StrConv(NormalizedName, vbProperCase)
        End If

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a blank name
        If Len(NormalizedName) = 0 Then
            Err.Raise vbObjectError + EX_ERR_BLANK_NAME, PROC_NAME, _
                "PersonName cannot be blank"
        End If
    'Reject a name longer than the supported maximum
        If Len(NormalizedName) > EX_MAX_NAME_LENGTH Then
            Err.Raise vbObjectError + EX_ERR_LONG_NAME, PROC_NAME, _
                "PersonName cannot exceed " & CStr(EX_MAX_NAME_LENGTH) & _
                " characters"
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Resolve the prefix in effect
        Prefix = M_Greeting_GetPrefix()
    'Join the prefix and the name
        M_Greeting_Build = Prefix & " " & NormalizedName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error
        Err.Raise ErrorNumber, PROC_NAME, "Greeting could not be built: " & _
            ErrorDescription

End Function

Private Function M_Greeting_GetPrefix() As String

'
'------------------------------------------------------------------------------
'                           RESOLVE GREETING PREFIX
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the greeting prefix in effect
'
' WHY THIS EXISTS
'   The default applies until a caller configures a different prefix
'
' INPUTS
'   None
'
' RETURNS
'   The configured prefix, or EX_DEFAULT_PREFIX when none is configured
'
' BEHAVIOR
'   Returns the stored prefix when it is not empty
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-01-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Use the default when nothing is configured
        If LenB(mEX_Prefix) = 0 Then
            M_Greeting_GetPrefix = EX_DEFAULT_PREFIX
        Else
            M_Greeting_GetPrefix = mEX_Prefix
        End If

End Function
