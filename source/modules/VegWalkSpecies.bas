Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegWalkSpecies
' Level:        Framework class
' Version:      1.03
'
' Description:  VegWalk species object related properties, events, functions & procedures for UI display
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 4/19/2016
' References:   -
' Revisions:    BLC - 4/19/2016 - 1.00 - initial version
'               BLC - 8/8/2016  - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.02 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.03 - SaveToDb() code cleanup
' =================================

'---------------------
' Declarations
'---------------------
Private m_Species As New Species 'CoverSpecies??

Private m_VegWalkID As Long
Private m_IsSeedling As Boolean

'---------------------
' Events
'---------------------
Public Event InvalidIsSeedling(Value As Boolean)

'-- base events (coverspecies)
Public Event InvalidVegPlotID(Value As String)
Public Event InvalidPercentCover(Value As Integer)

'-- base events (species) --
Public Event InvalidMasterPlantCode(Value As String)
Public Event InvalidLUCode(Value As String)
Public Event InvalidFamily(Value As String)
Public Event InvalidSpecies(Value As String)
Public Event InvalidCode(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let VegWalkID(Value As Long)
    m_VegWalkID = Value
End Property

Public Property Get VegWalkID() As Long
    VegWalkID = m_VegWalkID
End Property

Public Property Let IsSeedling(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_IsSeedling = Value
    Else
        RaiseEvent InvalidIsSeedling(Value)
    End If
End Property

Public Property Get IsSeedling() As Boolean
    IsSeedling = m_IsSeedling
End Property

' ---------------------------
' -- base class properties --
' ---------------------------
' NOTE: required since VBA does not support direct inheritance
'       or polymorphism like other OOP languages
' ---------------------------
' base class = Species
' ---------------------------
Public Property Let ID(Value As Long)
    m_Species.ID = Value
End Property

Public Property Get ID() As Long
    ID = m_Species.ID
End Property

Public Property Let MasterPlantCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.MasterPlantCode = Value
    Else
        RaiseEvent InvalidMasterPlantCode(Value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_Species.MasterPlantCode
End Property

Public Property Let COfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.COfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_Species.COfamily
End Property

Public Property Let UTfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.UTfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_Species.UTfamily
End Property

Public Property Let WYfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.WYfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_Species.WYfamily
End Property

Public Property Let COspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.COspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_Species.COspecies
End Property

Public Property Let UTspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.UTspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_Species.UTspecies
End Property

Public Property Let WYspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.WYspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_Species.WYspecies
End Property

Public Property Let LUCode(Value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(Value) And IsBetween(Len(Value), 1, 6, True) Then
        m_Species.LUCode = Value
    Else
        RaiseEvent InvalidLUCode(Value)
    End If
End Property

Public Property Get LUCode() As String
    LUCode = m_Species.LUCode
End Property

Public Property Let MasterFamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterFamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_Species.MasterFamily
End Property

Public Property Let MasterCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.MasterCode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_Species.MasterCode
End Property

Public Property Let MasterSpecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterSpecies = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_Species.MasterSpecies
End Property

Public Property Let UTcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.UTcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_Species.UTcode
End Property

Public Property Let COcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.COcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_Species.COcode
End Property

Public Property Let WYcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.WYcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_Species.WYcode
End Property

Public Property Let MasterCommonName(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterCommonName = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_Species.MasterCommonName
End Property

Public Property Let Lifeform(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Lifeform = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_Species.Lifeform
End Property

Public Property Let Duration(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Duration = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_Species.Duration
End Property

Public Property Let Nativity(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Nativity = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Nativity() As String
    Nativity = m_Species.Nativity
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========

' ---------------------------------
' SUB:          GetClass
' Description:  Retrieve a new instance of the class
'               --------------------------------------------------------------------------
'               Classes in a library with PublicNotCreateable instancing cannot
'               create items of the class in other projects (using the New keyword)
'               Variables can be declared, but the class object isn't created
'
'               This function allows other projects to create new instances of the class object
'               In referencing projects, set a reference to this project & call the GetClass()
'               function to create the new class object:
'                   Dim NewVegWalkSpecies as framework.VegWalkSpecies
'                   Set NewVegWalkSpecies = framework.GetClass()
'               --------------------------------------------------------------------------
' Assumptions:  -
' Parameters:   -
' Returns:      New instance of the class
' Throws:       none
' References:
'   Chip Pearson, November 6, 2013
'   http://www.cpearson.com/excel/classes.aspx
' Source/date:  -
' Adapted:      Bonnie Campbell, September 21, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/21/2016 - initial version
' ---------------------------------
Public Function GetClass() As VegWalkSpecies
On Error GoTo Err_Handler

    Set GetClass = New VegWalkSpecies

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetClass[VegWalkSpecies class])"
    End Select
    Resume Exit_Handler
End Function

'======== Standard Methods ==========

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    MsgBox "Initializing...", vbOKOnly
    
    Set m_Species = New Species

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[VegWalkSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
    
    MsgBox "Terminating...", vbOKOnly
        
    Set m_Species = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[VegWalkSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup veg walk species based on the lookup code
' Parameters:   luCode - species 6-character lookup code from NCPN master plants (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(LUCode As String)
On Error GoTo Err_Handler
    
            m_Species.Init (LUCode)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[VegWalkSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save veg walk species based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegwalk_species"
    
    Dim Params(0 To 5) As Variant
    
    With Me
        Params(0) = "VegWalkSpecies"
        Params(1) = .VegWalkID
        Params(2) = .MasterPlantCode
        Params(3) = .IsSeedling
        
        If IsUpdate Then
            Template = "u_vegwalk_species"
            Params(4) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
'    'add a record for created by
'    Dim act As New RecordAction
'
'    With act
'        .RefAction = "R"
'        .ContactID = TempVars("UserID")
'        .RefID = Me.ID
'        .RefTable = "VegWalkSpecies"
'        .SaveToDb
'    End With


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[VegWalkSpecies class])"
    End Select
    Resume Exit_Handler
End Sub