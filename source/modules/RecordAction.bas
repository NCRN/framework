Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        RecordAction
' Level:        Framework class
' Version:      1.06
'
' Description:  Record action object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 7/26/2016 - 1.01 - revised Action to RefAction to avoid conflict (Jet reserved word)
'               BLC - 8/8/2016  - 1.02 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.03 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.04 - SaveToDb() code cleanup
'               BLC - 10/6/2017 - 1.05 - removed GetClass() after Factory class instatiation implemented
'               BLC - 12/5/2017 - 1.06 - update to handle passed in ActionDate
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_RefAction As String
Private m_RefTable As String
Private m_RefID As Long
Private m_ContactID As Long
Private m_ActionType As String
Private m_ActionDate As Date

'---------------------
' Events
'---------------------
Public Event InvalidAction(Value As String)
Public Event InvalidRefTable(Value As String)
Public Event InvalidRefID(Value As Long)
Public Event InvalidContactID(Value As Long)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let RefTable(Value As String)
    If ValidateString(Value, "alphadashunderscore") Then
        m_RefTable = Value
    Else
        RaiseEvent InvalidRefTable(Value)
    End If
End Property

Public Property Get RefTable() As String
    RefTable = m_RefTable
End Property

Public Property Let RefID(Value As Long)
    m_RefID = Value
End Property

Public Property Get RefID() As Long
    RefID = m_RefID
End Property

Public Property Let ContactID(Value As Long)
    m_ContactID = Value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
End Property

'Action type is verbose for action
Public Property Let ActionType(Value As String)
    Select Case Value
        Case "Observe"
            Me.RefAction = "O"
        Case "Record"
            Me.RefAction = "R"
        Case "DataEntry"
            Me.RefAction = "DE"
        Case "Download"
            Me.RefAction = "D"
        Case "Upload"
            Me.RefAction = "U"
        Case "Change"
            Me.RefAction = "E"
        Case "Verify"
            Me.RefAction = "V"
        Case "Certify"
            Me.RefAction = "C"
    End Select

    m_ActionType = Value
End Property

Public Property Get ActionType() As String
    ActionType = m_ActionType
End Property

Public Property Let RefAction(Value As String)
    Dim aryActions() As String
    aryActions = Split(RECORD_ACTIONS, ",")
    
    If IsInArray(m_RefAction, aryActions) Then
        m_RefAction = Value
    Else
        RaiseEvent InvalidAction(Value)
    End If
End Property

Public Property Get RefAction() As String
    RefAction = m_RefAction
End Property

Public Property Let ActionDate(Value As Date)
    m_ActionDate = Value
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_ActionDate
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========
' handled by Factory class

'======== Standard Methods ==========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[RecordAction class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[RecordAction class])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save data to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 12/5/2017 - updated to handle passed in action date
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler

    Dim Params(0 To 4) As Variant

    Params(0) = Me.RefTable
    Params(1) = Me.RefID
    Params(2) = Me.ContactID
    Params(3) = Me.RefAction
    Params(4) = Ne(Me.ActionDate, CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM")))
 
    Me.ID = SetRecord("i_record_action", Params)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[RecordAction class])"
    End Select
    Resume Exit_Handler
End Sub