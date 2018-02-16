Attribute VB_Name = "StatePersistence"
Option Explicit
Public Sub SaveValidationData(state As LicensingAPI_COM.COMValidationState, _
    registeredUser As String, registeredEmail As String, path As String)
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(path, True)
    ts.WriteLine "Registered: " & CStr(state.Registered)
    ts.WriteLine "FailedRegistrationTimes: " & GetSeparatedArray(state.FailedRegistrationTimes)
    ts.WriteLine "HasRegisteredAt: " & CStr(state.HasRegisteredAt)
    If (state.HasRegisteredAt) Then
        ts.WriteLine "RegisteredAt: " & CStr(state.RegisteredAt)
    End If
    ts.WriteLine "FirstLaunchTime: " & CStr(state.FirstLaunchTime)
    ts.WriteLine "FailedValidationTimes: " & GetSeparatedArray(state.FailedValidationTimes)
    ts.WriteLine "HasLastSuccessfulValidationTime: " & CStr(state.HasLastSuccessfulValidationTime)
    If (state.HasLastSuccessfulValidationTime) Then
        ts.WriteLine "LastSuccessfulValidationTime: " & CStr(state.LastSuccessfulValidationTime)
    End If
    ts.WriteLine "LastValidationStatus: " & CStr(state.LastValidationStatus)
    ts.WriteLine "RegisteredToUser: " & registeredUser
    ts.WriteLine "RegisteredToEmail: " & registeredEmail
    ts.WriteLine "Features: " & GetSeparatedArray(state.Features)
    ts.WriteLine "InstanceId: " & state.InstanceId
    ts.WriteLine "GracePeriodForValidationFailures: " & CStr(state.GracePeriodForValidationFailures_seconds)
    ts.WriteLine "SessionDuration: " & CStr(state.SessionDuration_seconds)
    ts.WriteLine "HasUserId: " & CStr(state.HasUserId)
    If (state.HasUserId) Then
        ts.WriteLine "UserId: " & CStr(state.UserId)
    End If
    ts.WriteLine "ComponentsLoaded: " & CStr(state.ComponentsLoaded)
    If (state.ComponentsLoaded) Then
        ts.WriteLine "ComponentsCount: " & CStr(UBound(state.components))
        
        Dim component As COMComponentInfo
        Dim allComps() As COMComponentInfo
        allComps = state.components
        Dim i As Integer
        For i = 0 To UBound(state.components)
            Set component = allComps(i)
            ts.WriteLine vbTab & "Name: " & component.name
            ts.WriteLine vbTab & "LicenseModel: " & CStr(component.LicenseModel)
            ts.WriteLine vbTab & "TokensRequired: " & CStr(component.TokensRequired)
            ts.WriteLine vbTab & "HasFreeTrial: " & CStr(component.HasFreeTrial)
            ts.WriteLine vbTab & "FreeTrialExpirationTime: " & CStr(component.FreeTrialExpirationTime)
            ts.WriteLine vbTab & "OriginalTokensRequired: " & CStr(component.OriginalTokensRequired)
        Next
    End If
    ts.WriteLine "--"
    ts.Close
End Sub
Public Function LoadValidationData(path As String, ByRef registeredUser As String, _
    ByRef registeredEmail As String) As COMValidationState
        
    Dim ret As New COMValidationState
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(path, ForReading)
    
    ret.Registered = GetBoolean(ts.ReadLine, "Registered")
    ret.SetFailedRegistrationTimes GetDateArray(ts.ReadLine, "FailedRegistrationTimes")
    ret.HasRegisteredAt = GetBoolean(ts.ReadLine, "HasRegisteredAt")
    If (ret.HasRegisteredAt) Then
        ret.RegisteredAt = GetDate(ts.ReadLine, "RegisteredAt")
    End If
    ret.FirstLaunchTime = GetDate(ts.ReadLine, "FirstLaunchTime")
    ret.SetFailedValidationTimes GetDateArray(ts.ReadLine, "FailedValidationTimes")
    ret.HasLastSuccessfulValidationTime = GetBoolean(ts.ReadLine, "HasLastSuccessfulValidationTime")
    If (ret.HasLastSuccessfulValidationTime) Then
        ret.LastSuccessfulValidationTime = GetDate(ts.ReadLine, "LastSuccessfulValidationTime")
    End If
    ret.LastValidationStatus = GetInteger(ts.ReadLine, "LastValidationStatus")
    registeredUser = GetValue(ts.ReadLine, "RegisteredToUser")
    registeredEmail = GetValue(ts.ReadLine, "RegisteredToEmail")
    ret.SetFeatures GetStringArray(ts.ReadLine, "Features")
    ret.InstanceId = GetValue(ts.ReadLine, "InstanceId")
    ret.GracePeriodForValidationFailures_seconds = GetDouble(ts.ReadLine, "GracePeriodForValidationFailures")
    ret.SessionDuration_seconds = GetDouble(ts.ReadLine, "SessionDuration")
    ret.HasUserId = GetBoolean(ts.ReadLine, "HasUserId")
    If (ret.HasUserId) Then
        ret.UserId = GetValue(ts.ReadLine, "UserId")
    End If
    ret.ComponentsLoaded = GetBoolean(ts.ReadLine, "ComponentsLoaded")
    If (ret.ComponentsLoaded) Then
        Dim i, c As Integer
        Dim comps() As COMComponentInfo
        c = GetInteger(ts.ReadLine, "ComponentsCount")
        ReDim comps(0 To c - 1)
        For i = 1 To c
            Dim component As New COMComponentInfo
            component.name = GetValue(Trim(ts.ReadLine), "Name")
            component.LicenseModel = GetInteger(Trim(ts.ReadLine), "LicenseModel")
            component.TokensRequired = GetInteger(Trim(ts.ReadLine), "TokensRequired")
            component.HasFreeTrial = GetBoolean(Trim(ts.ReadLine), "HasFreeTrial")
            component.FreeTrialExpirationTime = GetDate(Trim(ts.ReadLine), "FreeTrialExpirationTime")
            component.OriginalTokensRequired = GetInteger(Trim(ts.ReadLine), "OriginalTokensREquired")
            Set comps(i - 1) = component
        Next
        ret.SetComponentsInfo comps
    End If
    ts.Close
    Set LoadValidationData = ret
End Function
Private Function GetSeparatedArray(coll As Variant) As String ' this can be optimized a lot!
    Dim final As String
    Dim s As Variant
    Dim i As Integer
    i = 0
    final = ""
    For Each s In coll
        If (i > 0) Then final = final & ";"
        final = final & CStr(s)
        i = i + 1
    Next
    GetSeparatedArray = final
End Function
Private Function GetValue(line As String, fieldName As String) As String
    GetValue = Right(line, Len(line) - Len(fieldName) - 2)
End Function
Private Function GetBoolean(line As String, fieldName As String) As Boolean
    Dim value As String
    value = GetValue(line, fieldName)
    If (value = "True") Then
        GetBoolean = True
    Else
        GetBoolean = False
    End If
End Function
Private Function GetInteger(line As String, fieldName As String) As Integer
    Dim value As String
    value = GetValue(line, fieldName)
    GetInteger = CInt(value)
End Function
Private Function GetDouble(line As String, fieldName As String) As Double
    Dim value As String
    value = GetValue(line, fieldName)
    GetDouble = CDbl(value)
End Function
Private Function GetDate(line As String, fieldName As String) As Date
    Dim value As String
    value = GetValue(line, fieldName)
    GetDate = CDate(value)
End Function
Private Function GetStringArray(line As String, fieldName As String) As String()
    Dim value As String
    Dim values() As String
    Dim ret() As String
    Dim i, length As Integer
    value = GetValue(line, fieldName)
    values = Split(value, ";")
    length = UBound(values) - LBound(values) + 1
    If (length = 0) Then
        GetStringArray = values
        Exit Function
    End If
    ReDim ret(0 To length - 1)
    length = 0
    For i = LBound(values) To UBound(values)
        If (values(i) <> vbNullString) Then
            ret(i - LBound(values)) = values(i)
            length = length + 1
        End If
    Next
    ReDim ret(0 To length - 1)
    GetStringArray = ret
End Function
Private Function GetDateArray(line As String, fieldName As String) As Date()
    Dim i, length As Integer
    Dim values() As String
    Dim ret() As Date
    values = GetStringArray(line, fieldName)
    If (UBound(values) - LBound(values)) <= 0 Then
        GetDateArray = ret
        Exit Function
    End If
    ReDim ret(0 To UBound(values))
    For i = 0 To UBound(values)
        ret(i) = CDate(values(i))
    Next
    GetDateArray = ret
End Function
