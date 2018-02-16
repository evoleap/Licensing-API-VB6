VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "elm VB6 Licensing API Tester"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerSession 
      Left            =   7080
      Top             =   4800
   End
   Begin RichTextLib.RichTextBox txtLog 
      Height          =   2295
      Left            =   0
      TabIndex        =   17
      Top             =   5160
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtProductGuid 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   "E85BB4BC-034A-49A5-88A3-C40D65F9E862"
      Top             =   540
      Width           =   5175
   End
   Begin VB.TextBox txtLicenseKey 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Text            =   "A1A9C-1M8C9-0UJOF-NTQF5-KHIUP"
      Top             =   120
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6376
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Register"
      TabPicture(0)   =   "frmMain.frx":007D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "btnRegister"
      Tab(0).Control(2)=   "txtInstanceId"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Session"
      TabPicture(1)   =   "frmMain.frx":0099
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "btnBeginSession"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnEndSession"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "btnComponent"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Components"
      TabPicture(2)   =   "frmMain.frx":00B5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "listComponents"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton btnComponent 
         Caption         =   "Update Components"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   2175
      End
      Begin MSComctlLib.ListView listComponents 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   30
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5106
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Component Details"
         Height          =   3135
         Left            =   -71640
         TabIndex        =   20
         Top             =   360
         Width           =   4095
         Begin VB.TextBox txtEntitlement 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   2280
            Width           =   2175
         End
         Begin VB.CommandButton btnCheckoutCheckin 
            Caption         =   "Check out"
            Height          =   375
            Left            =   1200
            TabIndex        =   29
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtFreeTrialExpiration 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.TextBox txtTokensRequired 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtComponentLicenseModel 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Entitlement"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Free trial expiration"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Under free trial?"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Tokens Required"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "License Model"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Session Details"
         Height          =   3015
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtSessionExpiry 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   16
            Text            =   "N/A"
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtSessionID 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   14
            Text            =   "N/A"
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtSessionActiveStatus 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   12
            Text            =   "No"
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "Session Expires:"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Session ID:"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Session Active?"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton btnEndSession 
         Caption         =   "End Session"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton btnBeginSession 
         Caption         =   "Begin Session"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtInstanceId 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72840
         TabIndex        =   7
         Text            =   "Not registered"
         Top             =   1020
         Width           =   5295
      End
      Begin VB.CommandButton btnRegister 
         Caption         =   "Register"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label label8 
         Caption         =   "Components"
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Instance Id:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Product Guid:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "License Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_licensingController As LicensingController
Attribute m_licensingController.VB_VarHelpID = -1
Private m_UserName As String
Private m_Email As String
Private m_selectedComponent As String


Private Sub btnBeginSession_Click()
    AddLog "Starting session"
    Call m_licensingController.BeginSession
End Sub

Private Sub btnCheckoutCheckin_Click()
    Dim components(0 To 0) As String
    components(0) = m_selectedComponent
    m_licensingController.CheckoutComponent components
End Sub

Private Sub btnComponent_Click()
    m_licensingController.GetComponentInfo
End Sub

Private Sub btnEndSession_Click()
    m_licensingController.EndSession
End Sub


Private Sub btnRegister_Click()
    m_licensingController.Register txtLicenseKey.text, m_UserName, m_Email
End Sub

Private Sub btnSave_Click()
    SaveValidationData m_licensingController.ValidationState, m_UserName, m_Email, "C:\state.txt"
End Sub

Private Sub Form_Load()
    m_UserName = "Test User"
    m_Email = "test@test.com"
    
    SetVersion 1, 0, 0, 0
    ReadPublicKeyFile "C:\LicensingAPITester\publickey.txt"
    SetProductId txtProductGuid.text
    
    Set m_licensingController = New LicensingController
    m_licensingController.Initialize "C://state.txt", timerSession, m_UserName, m_Email
End Sub
Private Sub AddLog(text As String)
    Dim existing As String
    existing = txtLog.text
    txtLog.text = existing & vbCrLf & Year(Now) & "-" & _
        Month(Now) & "-" & Day(Now) & " " & Hour(Now) & ":" & _
        Minute(Now) & ":" & Second(Now) & " - " & text
    
    
End Sub



Private Sub m_controlManager_OnRegistrationCompleted(ByVal result As LicensingAPI_COM.ICOMRegistrationResult)
    If (result.Success) Then
        AddLog ("Registration succeeded")
    Else
        AddLog ("Registration failed. " & result.ErrorMessage)
        txtInstanceId.text = "Not registered."
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    m_licensingController.EndSession
End Sub






Private Sub listComponents_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim compInfo As COMComponentInfo
    Dim arComponents() As COMComponentInfo
    arComponents = m_licensingController.ValidationState.components
    Set compInfo = arComponents(Item.Index - 1)
    m_selectedComponent = compInfo.name
    If (compInfo.LicenseModel = 1) Then
        txtComponentLicenseModel.text = "Token"
    Else
        txtComponentLicenseModel.text = "Session"
    End If
    txtTokensRequired.text = CStr(compInfo.TokensRequired) & "/" & CStr(compInfo.OriginalTokensRequired)
    If (compInfo.HasFreeTrial) Then
        Check1.value = 1
    Else
        Check1.value = 0
    End If
    If (compInfo.HasFreeTrial) Then
        txtFreeTrialExpiration.text = CStr(compInfo.FreeTrialExpirationTime)
    Else
        txtFreeTrialExpiration.text = "N/A"
    End If
    Dim entitlement As COMComponentEntitlementInfo
    Set entitlement = m_licensingController.GetEntitlementForComponent(Item.text)
    If (entitlement Is Nothing) Then
        txtEntitlement.text = "N/A"
    Else
        Dim strText As String
        strText = ""
        If (compInfo.LicenseModel = 1) Then
            strText = CStr(entitlement.TokenUsage.TokensEntitled) & " tokens"
            If (entitlement.TokenUsage.TokensInUse > 0) Then
                strText = strText & ". " & CStr(entitlement.TokenUsage.TokensInUse) & " in use"
            End If
            If (entitlement.TokenUsage.TokensInUseBySession > 0) Then
                strText = strText & " (" & CStr(entitlement.TokenUsage.TokensInUseBySession) & " used here)."
            Else
                strText = strText & "."
            End If
        Else
            strText = CStr(entitlement.SessionUsage.SessionsEntitled) & " sessions"
            If (entitlement.SessionUsage.SessionsInUse > 0) Then
                strText = strText & ". " & CStr(entitlement.SessionUsage.SessionsEntitled) & " in use"
            End If
            If (entitlement.SessionUsage.InUseBySession) Then
                strText = strText & " (1 used here)."
            Else
                strText = strText & "."
            End If
        End If
        txtEntitlement.text = strText
    End If
End Sub

Private Sub m_licensingController_OnInitialized()
    Call ChangeRegisterButtonState
    btnBeginSession.Enabled = True
End Sub

Private Sub m_licensingController_OnLogTextAdded(text As String)
    AddLog text
End Sub

Private Sub m_licensingController_OnRegistrationCompleted(result As LicensingAPI_COM.COMRegistrationResult)
    Call ChangeRegisterButtonState
    If Not (m_licensingController.ValidationState Is Nothing) Then
        If m_licensingController.ValidationState.Registered Then
            txtInstanceId.text = m_licensingController.ValidationState.InstanceId
        End If
    End If
End Sub

Private Sub m_licensingController_OnSessionUpdated(result As LicensingAPI_COM.COMSessionValidity)
    If (result.IsInUnregisteredGracePeriod) Then
        AddLog "Session in unregistered grace period"
    ElseIf (result.IsInValidationFailureGracePeriod) Then
        AddLog "Validation failed. Session in validation grace period. Will end in " & CStr(result.ValidityDuration_seconds) & " seconds"
    Else
        txtSessionActiveStatus.text = "Active"
        Dim SessionState As LicensingAPI_COM.ValidatedSessionState
        Set SessionState = m_licensingController.SessionState
        txtSessionID.text = SessionState.SessionKey
        txtSessionExpiry.text = CStr(result.ValidityDuration_seconds)
    End If
End Sub

Private Sub ChangeRegisterButtonState()
    If Not (m_licensingController.ValidationState Is Nothing) Then
        btnRegister.Enabled = Not (m_licensingController.ValidationState.Registered)
    Else
        btnRegister.Enabled = m_licensingController.Initialized
    End If
End Sub

Private Sub m_licensingController_OnStateChanged(state As LicensingAPI_COM.COMValidationState)
    txtInstanceId.text = state.InstanceId
    If state.Registered Then
        SSTab1.TabEnabled(0) = False
        If SSTab1.Tab = 0 Then SSTab1.Tab = 1
        listComponents.ListItems.Clear
        Dim arComp() As COMComponentInfo
        
        Dim i As Integer
        arComp = m_licensingController.ValidationState.components
        For i = 0 To UBound(arComp)
            listComponents.ListItems.Add i + 1, arComp(i).name, arComp(i).name
        Next i
        
    End If
End Sub

