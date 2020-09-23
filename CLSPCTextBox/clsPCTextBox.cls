VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPCTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private Const SW_SHOWNORMAL = 1

#If Win32 Then
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

#Else
    Private Declare Function ShellExecute Lib "shell.dll" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If



'Local copy of the Textbox object.  'WithEvents' key word will allow the class
'to handle the Textbox events.
Private WithEvents objTextBox As TextBox
Attribute objTextBox.VB_VarHelpID = -1

'String values for error handling
Private Const str_NotInitialized = "Object not properly initialized"
Private Const str_InvalidObject = "Invalid TextBox object"
Private Const str_ComFailure = "clsPCTextBox component failure"
        
'Variable to indicate if validation of Postcode is to be processed
Private m_VALIDATEPOSTCODE   As Boolean

'Error Handling
Private Const Err_BASE = 18320 + vbObjectError

Public Enum Err_clsPCTextBox
    err_NotInitialized = Err_BASE + 1
    err_InvalidObject
    err_ComFailure
End Enum

'Called to set up the initial values of the class
Public Sub Init(TextBoxobj As TextBox, Optional ValidatePC As Boolean = False)
    On Error GoTo errComponentFailure
    
    'ensure valid Textbox object
    If TextBoxobj Is Nothing Then
        On Error GoTo 0
        Err.Raise err_InvalidObject, App.EXEName & ".clsPCTextBox", str_InvalidObject
    End If
    
    Set objTextBox = TextBoxobj
    m_VALIDATEPOSTCODE = ValidatePC
Exit Sub

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Sub
  

Public Sub Get_Map_Location()
'  Created by * * * * * * * * * * * * * * * * * *
'  Jon Webb, Compass Computing  * * * * * * * * *
'   * * * * * * * * * * * * * * * * * * *  on  13/03/2001
    On Error GoTo errComponentFailure
    
    Dim iret As Long
    Dim URL As String
    URL = "http://uk.multimap.com/p/browse.cgi?pc=" & objTextBox.Text & "&title=MapLocation"
    
    iret = ShellExecute(objTextBox.hwnd, vbNullString, URL, _
            vbNullString, "c:\", SW_SHOWNORMAL)
            
    Exit Sub
         
errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Sub

'Get/Let the ValidatePostCode property
Public Property Get ValidatePostCode() As Boolean
    On Error GoTo errComponentFailure

    ValidatePostCode = m_VALIDATEPOSTCODE

Exit Property

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Property

Public Property Let ValidatePostCode(Value As Boolean)
    On Error GoTo errComponentFailure

    m_VALIDATEPOSTCODE = Value

Exit Property

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Property

'Get/Set the TextBox object. Used if init not called
Public Property Get TextBox() As TextBox
    On Error GoTo errComponentFailure

    Set TextBox = objTextBox

    Exit Property

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Property

Public Property Set TextBox(ByVal Value As TextBox)
    On Error GoTo errComponentFailure
    
    'ensure valid Textbox object
    If Value Is Nothing Then
        On Error GoTo 0
        Err.Raise err_InvalidObject, App.EXEName & ".clsPCTextBox", str_InvalidObject
    End If

    Set objTextBox = Value

    Exit Property

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Property

'Procedure to clear the contence of the Textbox Object
Public Sub Clear()
    On Error GoTo errComponentFailure
    
    'ensure valid Textbox object referenced within the class
    If objTextBox Is Nothing Then
        On Error GoTo 0
        Err.Raise err_NotInitialized, App.EXEName & ".clsPCTextBox", str_NotInitialized
    End If

    objTextBox.Text = vbNullString

Exit Sub

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Sub

'Destroy Textbox object
Private Sub Detach()
    On Error GoTo errComponentFailure

    If Not (objTextBox Is Nothing) Then Set objTextBox = Nothing

Exit Sub

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Sub

'Detect if the referenced object has lost focus
Private Sub objTextBox_LostFocus()
    On Error GoTo errComponentFailure
    
    'ensure valid Textbox object referenced within the class
    If objTextBox Is Nothing Then
        On Error GoTo 0
        Err.Raise err_NotInitialized, App.EXEName & ".clsPCTextBox", str_NotInitialized
    End If
    
    'Check to see if Textbox contains a value
    If Len(objTextBox) = 0 Then GoTo Line_Bypass
    
    'Check to see if postcode validation required
    If Not m_VALIDATEPOSTCODE Then GoTo Line_Bypass
    
    'If validation failed then setfocus to the Textbox Object
    If Not Validate_PCode_Format Then objTextBox.SetFocus

Line_Bypass:
Exit Sub

errComponentFailure:
    Err.Raise err_ComFailure, App.EXEName & ".clsPCTextBox", str_ComFailure
End Sub

Private Function Validate_PCode_Format() As Boolean
'=================================================
    
    '  Created by * * * * * * * * * * * * * * * * * *
    '  Nicholas Andrews * * * * * * * * * * * * * * *
    '  on  17/03/2003 * * * * * * * * * * * * * * * *
    
'=================================================

    Dim strPostCode As String 'Temp string to hold initial value of the Textbox Object
    Dim strInward As String 'Used during the insert of a Space within the postcode
    Dim lngPostCode As Long 'Used for holding the length of strPostCode
    Dim strCHRInvalid As String
    Dim lngCHRPos As Long
    Dim i As Long
    
    On Error GoTo Validate_PCode_Format_ErrHandler
    
    'Set the temp string value to equal the current contender of the
    'Textbox Object with ALL spaces removed.
    'As we all know, end users are idiots at the best of times.
    'I have seen a number of spaces in a variety of
    'Locations within the postcode entry. Because the space is always
    'The 4th character it is far quicker to remove all first and insert
    'One later, than it is to account for all permutations end user incompetents

    strPostCode = Replace$(objTextBox.Text, " ", vbNullString, 1)
    
    'With all spaces removed, a valid postcode will be between 5 & 7 characters
    'Fail validation if strPostCode does not meet these requirements
    lngPostCode = Len(strPostCode)
    If lngPostCode <= 4 Then GoTo Validate_Fail
    If lngPostCode > 7 Then GoTo Validate_Fail
    
    'Let strInward = the last 3 characters of strPostCode
    strInward = Right$(strPostCode, 3)
    
    'Check for invalid characters (C I K M O V)in the inward section.
    For i = 1 To 6
        Select Case i
        Case 1: strCHRInvalid = "c"
        Case 2: strCHRInvalid = "i"
        Case 3: strCHRInvalid = "k"
        Case 4: strCHRInvalid = "m"
        Case 5: strCHRInvalid = "o"
        Case 6: strCHRInvalid = "v"
        End Select
        lngCHRPos = InStr(1, strInward, strCHRInvalid, vbTextCompare)
        If lngCHRPos > 0 Then GoTo Validate_Fail
    Next
        
    'Then let strPostcode = the remaining characters by using the mid function.
    'Attach the space followed by strInward which holds the end string
    strPostCode = Mid$(strPostCode, 1, (lngPostCode - 3)) & " " & strInward
    
    'Re-assign the length value
    lngPostCode = Len(strPostCode)
    
    'Force the string value to uppercase
    strPostCode = UCase(strPostCode)
    
    'The select case statement will determine which format strPostCode is to
    'Be compared against relevant to its length.
    
    'The 'Like' function is used to validate strPostcode conforms to a valid format
    'If Valid then bypass the fail message
    Select Case lngPostCode
    Case 6
        If strPostCode Like "[A-Z][0-9] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][0-9] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
    Case 7
        If strPostCode Like "[A-Z][0-9][0-9] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][0-9][0-9] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][A-Z][0-9] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][A-Z][0-9] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][0-9][A-Z] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][0-9][A-Z] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
    Case 8
        If strPostCode Like "[A-Z][A-Z][0-9][0-9] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][A-Z][0-9][0-9] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][A-Z][0-9][A-Z] [0-9][A-Z][A-Z]" Then GoTo Validate_Pass
        If strPostCode Like "[A-Z][A-Z][0-9][A-Z] [A-Z][A-Z][A-Z]" Then GoTo Validate_Pass
    End Select

'if you are here then the validation process has failed
Validate_Fail:
    Dim msg As String
    'Display a nice message to the end user.  Informing them of the error
    'And a list of valid formats to aid them correcting their mistake.
    
    'Deleting every record in the database would probably be a better way
    'Of making sure they pay more attention in the future.:-)

    msg = "Postcode entered does not conform to one of the" & vbCr
    msg = msg & "following formats:" & vbCr & vbCr
    msg = msg & "Left side                                    Right side" & vbCr
    msg = msg & "AN                                            NAA" & vbCr
    msg = msg & "ANN                                          AAA" & vbCr
    msg = msg & "AAN" & vbCr & "ANA" & vbCr & "AANN" & vbCr & "AANA" & vbCr
    msg = msg & vbCr & "A = [A-Z]           N = [0-9]" & vbCr & vbCr
    msg = msg & "A space ALWAYS separates left and right sides." & vbCr & vbCr
    msg = msg & "C I K M O V are invalid characters for right side" & vbCr & "of the postcode" & vbCr
    MsgBox msg, vbInformation, "Format Error"
    
    'If fail skip the 'Pass' part of the function
    GoTo Validate_PCode_Format_Done

Validate_Pass:
        'Because we may have changed the initial string due to 'Case' and or 'Spaces'
        'Make TextBox object = strPostCode
        objTextBox.Text = strPostCode
        Validate_PCode_Format = True

Validate_PCode_Format_Done:
        Exit Function

Validate_PCode_Format_ErrHandler:
        Validate_PCode_Format = False
        Resume Validate_PCode_Format_Done
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    
    Call Detach
End Sub
