VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LicensingController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents m_sessionExtentionTimer As Timer
Attribute m_sessionExtentionTimer.VB_VarHelpID = -1
Private m_timeRemaining As Long
Private WithEvents m_licenseCheckTimer As Timer
Attribute m_licenseCheckTimer.VB_VarHelpID = -1


Private m_State As LicensingAPI_COM.COMValidationState
Private m_registeredToUser As String
Private m_registeredToEmail As String
Private m_userIdentity As LicensingAPI_COM.COMUserIdentity
Private m_instanceIdentity As LicensingAPI_COM.COMInstanceIdentity

Private WithEvents m_identityHelper As LicensingAPI_COM.HardwareBasedInstanceIdentityHelper
Attribute m_identityHelper.VB_VarHelpID = -1
Private m_instanceIdHelperInterface As LicensingAPI_COM.ICOMHardwareBasedInstanceIdentityHelper
Private WithEvents m_userIdentityHelper As LicensingAPI_COM.ActiveDirectoryBasedUserIdentityHelper
Attribute m_userIdentityHelper.VB_VarHelpID = -1
Private m_userIdHelperInterface As LicensingAPI_COM.ICOMActiveDirectoryBasedUserIdentityHelper
Private WithEvents m_controlManager As LicensingAPI_COM.COMControlManager
Attribute m_controlManager.VB_VarHelpID = -1
Private m_controlManagerInterface As LicensingAPI_COM.ICOMControlManager
Private m_sessionValidity As COMSessionValidity
Private m_compEntitlements() As COMComponentEntitlementInfo
Private m_filePath As String
Private m_initialized As Boolean
Public Event OnInitialized()
Public Event OnStateChanged(state As COMValidationState)
Public Event OnRegistrationCompleted(result As COMRegistrationResult)
Public Event OnSessionUpdated(result As COMSessionValidity)
Public Event OnLogTextAdded(text As String)
Public Property Get Initialized() As Boolean
    Initialized = m_initialized
End Property
Public Property Get ValidationState() As COMValidationState
    Set ValidationState = m_State
End Property
Public Property Get SessionState() As ValidatedSessionState
    Set SessionState = m_controlManagerInterface.SessionState
End Property
''''''''''''''''''''''''''''''''''''''''''''''
''''''''' PUBLIC METHODS '''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialize(savedStatePath As String, sessionTimer As Timer, _
    Optional userName As String = "", Optional userEmail As String = "")
    Dim fso As New FileSystemObject
    Set m_State = Nothing
    m_registeredToEmail = userEmail
    m_registeredToUser = userName
    m_filePath = savedStatePath
    If fso.FileExists(savedStatePath) Then
        Set m_State = LoadValidationData(savedStatePath, m_registeredToUser, m_registeredToEmail)
        Call RaiseStateChangedEvent
    End If
    Set m_identityHelper = New LicensingAPI_COM.HardwareBasedInstanceIdentityHelper
    Set m_instanceIdHelperInterface = m_identityHelper
    Set m_userIdentityHelper = New LicensingAPI_COM.ActiveDirectoryBasedUserIdentityHelper
    Set m_userIdHelperInterface = m_userIdentityHelper
        
    Dim lqs As New LicensingAPI_COM.LicensingQuickStart
    lqs.Host = "http://hou-pvp2-pwe10:3000"
    lqs.Timeout_seconds = 300#

    Dim ii As COMInstanceIdentity
    Set ii = New COMInstanceIdentity
    m_instanceIdHelperInterface.BeginAddAllHardwareKeys ii
    m_userIdHelperInterface.BeginGetIdentity
    
    Set m_sessionExtentionTimer = sessionTimer
    m_sessionExtentionTimer.Enabled = False
End Sub
Public Sub Register(licenseKey As String, userName As String, userEmail As String)
    Dim userInfo As New COMUserInfo
    userInfo.name = userName
    userInfo.Email = userEmail
    m_controlManagerInterface.Register licenseKey, userInfo
End Sub
Public Sub BeginSession()
    Call ValidateSessionCore
End Sub
Public Sub EndSession()
    If Not m_initialized Then Exit Sub
    If (m_sessionValidity Is Nothing) Then Exit Sub
    Call m_controlManagerInterface.EndSession
End Sub
Public Sub GetComponentInfo()
    If Not m_initialized Then Exit Sub
    If (m_sessionValidity Is Nothing) Then Exit Sub
    RaiseEvent OnLogTextAdded("Getting component information")
    Call m_controlManagerInterface.GetComponentsStatus
End Sub
Public Sub CheckoutComponent(name() As String)
    If Not m_initialized Then Exit Sub
    If (m_sessionValidity Is Nothing) Then Exit Sub
    RaiseEvent OnLogTextAdded("Checking out " & name(0))
    m_controlManagerInterface.CheckOutComponents name
End Sub
Public Function GetEntitlementForComponent(name As String) As COMComponentEntitlementInfo
    If (Not (IsArray(m_compEntitlements)) Or UBound(m_compEntitlements) < 0) Then
        Set GetEntitlementForComponent = Nothing
        Exit Function
    End If
    Dim i As Integer
    For i = 0 To UBound(m_compEntitlements)
        Dim comps() As String
        Dim entitlement As COMComponentEntitlementInfo
        Set entitlement = m_compEntitlements(i)
        comps = entitlement.components
        Dim j As Integer
        For j = 0 To UBound(comps)
            If (comps(j) = name) Then
                Set GetEntitlementForComponent = entitlement
                Exit Function
            End If
        Next j
    Next i
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''
''''''''' PRIVATE METHODS ''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''
''''''''' UTILITY CALLS  '''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateSessionCore()
    If Not m_initialized Then Exit Sub
    m_controlManagerInterface.ValidateSession 60#
End Sub
''''''''''''''''''''''''''''''''''''''''''''''
''''''''' EVENT HANDLERS '''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''

Private Sub m_controlManager_OnCheckoutComponentsCompleted(ByVal result As LicensingAPI_COM.ICOMComponentCheckoutResult)
    If (result.Success) Then
        RaiseEvent OnLogTextAdded("Component checkout successful")
    Else
        RaiseEvent OnLogTextAdded("Component checkout failed. " & result.ErrorMessage)
    End If
    Call UpdateStateFromControlManager

End Sub

Private Sub m_controlManager_OnGetComponentsStatusCompleted(ByVal result As LicensingAPI_COM.ICOMComponentsStatus)
    RaiseEvent OnLogTextAdded("Component information obtained")
    Call UpdateStateFromControlManager
    m_compEntitlements = result.ComponentEntitlements
End Sub

Private Sub m_controlManager_OnRegistrationCompleted(ByVal result As LicensingAPI_COM.ICOMRegistrationResult)
    RaiseEvent OnRegistrationCompleted(result)
    RaiseEvent OnLogTextAdded("Registration complete")
    Call UpdateStateFromControlManager
End Sub

Private Sub m_controlManager_OnValidateSessionCompleted(ByVal result As LicensingAPI_COM.ICOMSessionValidity)
    RaiseEvent OnSessionUpdated(result)
    RaiseEvent OnLogTextAdded("Session updated")
    Set m_sessionValidity = result
    m_timeRemaining = CLng(result.ValidityDuration_seconds * 1000)
    Call SetTimerInterval
    Call UpdateStateFromControlManager
End Sub

Private Sub SetTimerInterval()
    Dim timeToSet As Long
    m_sessionExtentionTimer.Enabled = False
    If (m_timeRemaining > 60000) Then
        timeToSet = 60000
    Else
        timeToSet = m_timeRemaining
    End If
    m_sessionExtentionTimer.Interval = timeToSet
    m_timeRemaining = m_timeRemaining - timeToSet
    m_sessionExtentionTimer.Enabled = True
End Sub
Private Sub m_identityHelper_OnInstanceIdentityGathered(ByVal identity As LicensingAPI_COM.ICOMInstanceIdentity)
    Set m_instanceIdentity = identity
    RaiseEvent OnLogTextAdded("Instance identity built")
    Call CheckAndRaiseInitializedEvent
End Sub

Private Sub m_sessionExtentionTimer_Timer()
    If (m_timeRemaining <= 1) Then
        Set m_sessionValidity = Nothing
        RaiseEvent OnLogTextAdded("Initiating session extension")
        Call ValidateSessionCore
    Else
        Call SetTimerInterval
    End If
End Sub

Private Sub m_userIdentityHelper_OnUserIdentityGathered(ByVal identity As LicensingAPI_COM.ICOMUserIdentity)
    Set m_userIdentity = identity
    RaiseEvent OnLogTextAdded("User identity built")
    Call CheckAndRaiseInitializedEvent
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
''''''''' EVENT TRIGGERS '''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckAndRaiseInitializedEvent()
    If (m_userIdentity Is Nothing) Then Exit Sub
    If (m_instanceIdentity Is Nothing) Then Exit Sub
    Set m_controlManager = CreateManager(m_userIdentity, m_instanceIdentity, m_State, New COMControlStrategy)
    Set m_controlManagerInterface = m_controlManager
    m_initialized = True
    RaiseEvent OnInitialized
    RaiseEvent OnLogTextAdded("Licensing controller initialization complete")
End Sub

Private Sub RaiseStateChangedEvent()
    RaiseEvent OnStateChanged(m_State)
End Sub
Private Sub UpdateStateFromControlManager()
    Set m_State = m_controlManagerInterface.ValidationState
    SaveValidationData m_State, m_registeredToUser, m_registeredToEmail, m_filePath
    Call RaiseStateChangedEvent
End Sub
