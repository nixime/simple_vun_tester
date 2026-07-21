Option Explicit

' ==============================================================================
' PUBLIC METHOD 1: Calculate Base Finding Subscore
' Takes string values for Technical Impact, Acquired Privilege, 
' Acquired Privilege Layer, Internal Control Effectiveness, and Finding Confidence.
' ==============================================================================
Public Function CalculateBaseFinding(ti As String, ap As String, al As String, ic As String, fc As String) As Double
    Dim vTI As Double, vAP As Double, vAL As Double, vIC As Double, vFC As Double
    Dim fTI As Double
    
    vTI = GetVal_Impact(ti)
    vAP = GetVal_Privilege(ap)
    vAL = GetVal_Layer(al)
    vIC = GetVal_Control(ic)
    vFC = GetVal_Confidence(fc)
    
    If vTI > 0 Then fTI = 1 Else fTI = 0
    
    CalculateBaseFinding = ((10 * vTI + 5 * (vAP + vAL) + 5 * vFC) * fTI * vIC) * 4.0
    If CalculateBaseFinding > 100 Then CalculateBaseFinding = 100
    If CalculateBaseFinding < 0 Then CalculateBaseFinding = 0
End Function

' ==============================================================================
' PUBLIC METHOD 2: Calculate Attack Surface Subscore
' Takes string values for Required Privilege, Required Privilege Layer, 
' Access Vector, Authentication Strength, Level of Interaction, and Deployment Scope.
' ==============================================================================
Public Function CalculateAttackSurface(rp As String, rl As String, av As String, asStr As String, inStr As String, sc As String) As Double
    Dim vRP As Double, vRL As Double, vAV As Double, vAS As Double, vIN As Double, vSC As Double
    
    vRP = GetVal_Privilege(rp)
    vRL = GetVal_Layer(rl)
    vAV = GetVal_AccessVector(av)
    vAS = GetVal_AuthStrength(asStr)
    vIN = GetVal_Interaction(inStr)
    vSC = GetVal_DeploymentScope(sc)
    
    CalculateAttackSurface = (20 * (vRP + vRL + vAV) + 20 * vSC + 15 * vIN + 5 * vAS) / 100.0
    If CalculateAttackSurface > 1 Then CalculateAttackSurface = 1
    If CalculateAttackSurface < 0 Then CalculateAttackSurface = 0
End Function

' ==============================================================================
' PUBLIC METHOD 3: Calculate Environmental Subscore
' Takes string values for Business Impact, Likelihood of Discovery, 
' Likelihood of Exploit, External Control Effectiveness, and Prevalence.
' ==============================================================================
Public Function CalculateEnvironmental(bi As String, di As String, ex As String, ec As String, p As String) As Double
    Dim vBI As Double, vDI As Double, vEX As Double, vEC As Double, vP As Double
    Dim fBI As Double
    
    vBI = GetVal_Impact(bi)
    vDI = GetVal_Discovery(di)
    vEX = GetVal_Exploit(ex)
    vEC = GetVal_Control(ec)
    vP = GetVal_Prevalence(p)
    
    If vBI > 0 Then fBI = 1 Else fBI = 0
    
    CalculateEnvironmental = ((10 * vBI + 3 * vDI + 4 * vEX + 3 * vP) * fBI * vEC) / 20.0
    If CalculateEnvironmental > 1 Then CalculateEnvironmental = 1
    If CalculateEnvironmental < 0 Then CalculateEnvironmental = 0
End Function

' ==============================================================================
' PRIVATE MAPPING / LOOKUP FUNCTIONS BASED ON REFERENCE TABLES
' ==============================================================================

Private Function GetVal_Impact(s As String) As Double
    Select Case UCase(Trim(s))
        Case "C", "CRITICAL": GetVal_Impact = 1.0
        Case "H", "HIGH": GetVal_Impact = 0.75
        Case "M", "MEDIUM": GetVal_Impact = 0.5
        Case "L", "LOW": GetVal_Impact = 0.25
        Case "N", "NONE": GetVal_Impact = 0.0
        Case Else: GetVal_Impact = 0.5
    End Select
End Function

Private Function GetVal_Privilege(s As String) As Double
    Select Case UCase(Trim(s))
        Case "A", "ADMINISTRATOR": GetVal_Privilege = 0.0
        Case "R", "REGULAR USER": GetVal_Privilege = 0.5
        Case "N", "NONE": GetVal_Privilege = 1.0
        Case Else: GetVal_Privilege = 0.5
    End Select
End Function

Private Function GetVal_Layer(s As String) As Double
    Select Case UCase(Trim(s))
        Case "E", "ENTERPRISE": GetVal_Layer = 1.0
        Case "I", "INFRASTRUCTURE": GetVal_Layer = 0.8
        Case "S", "SYSTEM": GetVal_Layer = 0.6
        Case "A", "APPLICATION": GetVal_Layer = 0.4
        Case "N", "NETWORK": GetVal_Layer = 0.2
        Case Else: GetVal_Layer = 0.5
    End Select
End Function

Private Function GetVal_Control(s As String) As Double
    Select Case UCase(Trim(s))
        Case "N", "NONE": GetVal_Control = 1.0
        Case "L", "LIMITED / STANDARD", "LIMITED": GetVal_Control = 0.75
        Case "O", "OPPORTUNISTIC", "MODERATE": GetVal_Control = 0.5
        Case "W", "WEAK": GetVal_Control = 0.25
        Case "S", "STRONG": GetVal_Control = 0.25
        Case "H", "HIGH": GetVal_Control = 0.0
        Case Else: GetVal_Control = 1.0
    End Select
End Function

Private Function GetVal_Confidence(s As String) As Double
    Select Case UCase(Trim(s))
        Case "PR", "PROVEN": GetVal_Confidence = 1.0
        Case "T", "TRUE": GetVal_Confidence = 0.75
        Case "PL", "PROVEN LOCALLY": GetVal_Confidence = 0.5
        Case "PO", "POTENTIAL": GetVal_Confidence = 0.25
        Case "PF", "PROVEN FALSE": GetVal_Confidence = 0.0
        Case Else: GetVal_Confidence = 0.75
    End Select
End Function

Private Function GetVal_AccessVector(s As String) As Double
    Select Case UCase(Trim(s))
        Case "I", "INTERNET": GetVal_AccessVector = 1.0
        Case "W", "INTRANET": GetVal_AccessVector = 0.75
        Case "P", "PRIVATE NETWORK": GetVal_AccessVector = 0.5
        Case "L", "LOCAL": GetVal_AccessVector = 0.25
        Case "X", "PHYSICAL": GetVal_AccessVector = 0.0
        Case Else: GetVal_AccessVector = 0.5
    End Select
End Function

Private Function GetVal_AuthStrength(s As String) As Double
    Select Case UCase(Trim(s))
        Case "N", "NONE": GetVal_AuthStrength = 1.0
        Case "W", "WEAK": GetVal_AuthStrength = 0.66
        Case "S", "STRONG": GetVal_AuthStrength = 0.33
        Case Else: GetVal_AuthStrength = 0.0
    End Select
End Function

Private Function GetVal_Interaction(s As String) As Double
    Select Case UCase(Trim(s))
        Case "N", "NONE": GetVal_Interaction = 0.0
        Case "L", "LIMITED / STANDARD": GetVal_Interaction = 0.33
        Case "O", "OPPORTUNISTIC": GetVal_Interaction = 0.66
        Case "H", "HIGH", "ACTIVE": GetVal_Interaction = 1.0
        Case Else: GetVal_Interaction = 0.0
    End Select
End Function

Private Function GetVal_DeploymentScope(s As String) As Double
    Select Case UCase(Trim(s))
        Case "A", "ALL": GetVal_DeploymentScope = 1.0
        Case "C", "COMMON": GetVal_DeploymentScope = 0.75
        Case "O", "OCCASIONAL": GetVal_DeploymentScope = 0.5
        Case "R", "RARE": GetVal_DeploymentScope = 0.25
        Case "N", "NONE": GetVal_DeploymentScope = 0.0
        Case Else: GetVal_DeploymentScope = 0.5
    End Select
End Function

Private Function GetVal_Discovery(s As String) As Double
    Select Case UCase(Trim(s))
        Case "H", "HIGH": GetVal_Discovery = 1.0
        Case "M", "MEDIUM": GetVal_Discovery = 0.66
        Case "L", "LOW": GetVal_Discovery = 0.33
        Case "N", "NONE": GetVal_Discovery = 0.0
        Case Else: GetVal_Discovery = 0.5
    End Select
End Function

Private Function GetVal_Exploit(s As String) As Double
    Select Case UCase(Trim(s))
        Case "A", "ACTIVE": GetVal_Exploit = 1.0
        Case "P", "PROOF OF CONCEPT": GetVal_Exploit = 0.75
        Case "U", "THEORETICAL": GetVal_Exploit = 0.5
        Case "N", "NONE": GetVal_Exploit = 0.0
        Case Else: GetVal_Exploit = 0.0
    End Select
End Function

Private Function GetVal_Prevalence(s As String) As Double
    Select Case UCase(Trim(s))
        Case "W", "WIDESPREAD": GetVal_Prevalence = 1.0
        Case "C", "COMMON": GetVal_Prevalence = 0.75
        Case "O", "OCCASIONAL": GetVal_Prevalence = 0.5
        Case "R", "RARE": GetVal_Prevalence = 0.25
        Case "N", "NONE": GetVal_Prevalence = 0.0
        Case Else: GetVal_Prevalence = 0.5
    End Select
End Function