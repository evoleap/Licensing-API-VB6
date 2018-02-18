Attribute VB_Name = "ControlManagerFactory"
    Private s_Major, s_Minor, s_Revision, s_Build As Integer
Private s_ProductId As String
Private s_PublicKeyText As String

Public Function CreateManager(userIdentity As COMUserIdentity, instanceIdentity As COMInstanceIdentity, _
    savedState As COMValidationState, strategy As COMControlStrategy) As COMControlManager
    Dim cmf As New ControlManagerFactory
    Dim cmip As New ControlManagerInitializationParameters
    Dim cm As ICOMControlManager
    Dim userInfo As New COMUserInfo
    Set cmip.ControlStrategy = strategy
    Set cmip.instanceIdentity = instanceIdentity
    Set cmip.userIdentity = userIdentity
    Set cmip.savedState = savedState
    cmip.productId = s_ProductId
    cmip.PublicKeyText = s_PublicKeyText
    cmip.SetVersion s_Major, s_Minor, s_Revision, s_Build
    
    Set CreateManager = cmf.GetManager(cmip)

End Function
Public Sub SetVersion(major As Integer, minor As Integer, Optional build As Integer = 0, Optional rev As Integer = 0)
    s_Major = major
    s_Minor = minor
    s_Build = build
    s_Revision = rev
End Sub

Public Sub ReadPublicKeyFile(path As String)
    Dim fso As New Scripting.FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(path, ForReading)
    SetPublicKeyText ts.ReadAll
    ts.Close
End Sub
Public Sub SetPublicKeyText(keyText As String)
    s_PublicKeyText = keyText
End Sub
Public Sub SetProductId(id As String)
    s_ProductId = id
End Sub
