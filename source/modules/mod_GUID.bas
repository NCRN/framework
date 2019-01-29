Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_GUID
' Level:        Framework module
' Version:      1.01
' Description:  Framework-wide GUID related values, functions & subroutines
'
' References:
'   Ben Baird, unknown (downloaded 12/22/2005)
' Source/date:  John R. Boetsch, May 26, 2006
' Adapted:      Bonnie Campbell, January 2019
' Revisions:    JRB, 5/26/2006 - 1.00 - initial version
'                                       documentation & minimal edits
'               BLC, 1/23/2019 - 1.01 - moved & renamed from basUtilities (ForestVeg)
' =================================

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, September 30, 2016 for NCPN tools
' Adapted:      -
' Revisions:    BLC, 9/30/2016 - initial version
' ---------------------------------

'-----------------------------------------------------------------------
' Constants
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------
Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

' ---------------------------------
'  Properties
' ---------------------------------

'-----------------------------------------------------------------------
' Functions
'-----------------------------------------------------------------------
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long

Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" _
    (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' Sub:          CreateGUID
' Description:  Generates 16-byte globally-unique identifiers (GUIDs)
' Assumptions:  -
' Parameters:   -
' Returns:      a formatted guid as a string type, which can be saved directly
'               to either a string or a GUID field
' Throws:       none
' Requires:     CoCreateGuid API to generate GUID
'               StringFromGUID2 API to format as a string
' References:
'   Ben Baird, unknown (downloaded 12/22/2005)
' Source/date:  John R. Boetsch, May 26, 2006
' Adapted:      Bonnie Campbell, January 2019
' Revisions:    JRB, 5/26/2006 - 1.00 - initial version
'                                       documentation & minimal edits
'               BLC, 1/23/2019 - 1.01 - moved from basUtilities (ForestVeg),
'                                       renamed fxnGUIDGen >> CreateGUID
' ---------------------------------
Public Function CreateGUID()
On Error GoTo Err_Handler

    Dim uGUID As GUID       ' structured guid
    Dim sGUID As String     ' for storing the results
    Dim bGUID() As Byte     ' formatted string
    Dim lLen As Long
    Dim RetVal As Long
    lLen = 40
    bGUID = String(lLen, 0)

    ' use the API to generate the guid
    CoCreateGuid uGUID

    ' use the API to format as string
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    
    ' truncate the string
    CreateGUID = Left$(sGUID, RetVal)

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateGUID[mod_GUID])"
    End Select
    Resume Exit_Handler
End Function