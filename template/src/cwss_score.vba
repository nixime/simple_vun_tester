Option Explicit

Function ExtractCode(inputStr As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(inputStr, "(")
    endPos = InStr(inputStr, ")")
    
    If startPos > 0 And endPos > startPos Then
        ExtractCode = Trim(Mid(inputStr, startPos + 1, endPos - startPos - 1))
    Else
        ExtractCode = ""
    End If
End Function

' ==============================================================================
' Calculate Base Finding Subscore
' ==============================================================================
Public Function CalculateBaseFinding(ti As String, ap As String, al As String, ic As String, fc As String) As Double
    Dim vTI As Double, vAP As Double, vAL As Double, vIC As Double, vFC As Double
    Dim fTI As Double
    
    vTI = GetVal_TI(ExtractCode(ti))
    vAP = GetVal_AP(ExtractCode(ap))
    vAL = GetVal_AL(ExtractCode(al))
    vIC = GetVal_IC(ExtractCode(ic))
    vFC = GetVal_FC(ExtractCode(fc))
    
    If vTI > 0 Then fTI = 1 Else fTI = 0
    
    CalculateBaseFinding = ((10 * vTI + 5 * (vAP + vAL) + 5 * vFC) * fTI * vIC) * 4.0
    If CalculateBaseFinding > 100 Then CalculateBaseFinding = 100
    If CalculateBaseFinding < 0 Then CalculateBaseFinding = 0
End Function

' ==============================================================================
' Calculate Attack Surface Subscore
' ==============================================================================
Public Function CalculateAttackSurface(rp As String, rl As String, av As String, asStr As String, inStr As String, sc As String) As Double
    Dim vRP As Double, vRL As Double, vAV As Double, vAS As Double, vIN As Double, vSC As Double
    
    vRP = GetVal_RP(ExtractCode(rp))
    vRL = GetVal_RL(ExtractCode(rl))
    vAV = GetVal_AV(ExtractCode(av))
    vAS = GetVal_AS(ExtractCode(asStr))
    vIN = GetVal_IN(ExtractCode(inStr))
    vSC = GetVal_SC(ExtractCode(sc))
    
    CalculateAttackSurface = (20 * (vRP + vRL + vAV) + 20 * vSC + 15 * vIN + 5 * vAS) / 100.0
    If CalculateAttackSurface > 1 Then CalculateAttackSurface = 1
    If CalculateAttackSurface < 0 Then CalculateAttackSurface = 0
End Function

' ==============================================================================
' Calculate Environmental Subscore
' ==============================================================================
Public Function CalculateEnvironmental(bi As String, di As String, ex As String, ec As String, p As String) As Double
    Dim vBI As Double, vDI As Double, vEX As Double, vEC As Double, vP As Double
    Dim fBI As Double
    
    vBI = GetVal_BI(ExtractCode(bi))
    vDI = GetVal_DI(ExtractCode(di))
    vEX = GetVal_EX(ExtractCode(ex))
    vEC = GetVal_EC(ExtractCode(ec))
    vP = GetVal_P(ExtractCode(p))
    
    If vBI > 0 Then fBI = 1 Else fBI = 0
    
    CalculateEnvironmental = ((10 * vBI + 3 * vDI + 4 * vEX + 3 * vP) * fBI * vEC) / 20.0
    If CalculateEnvironmental > 1 Then CalculateEnvironmental = 1
    If CalculateEnvironmental < 0 Then CalculateEnvironmental = 0
End Function


' =====================================================================
' Complete CWSS v1.0.1 Component Weight Mapping Functions
' Naming Convention: GetVal_XX
' Returns official MITRE CWSS numeric weights for all metric factors
' =====================================================================

' ---------------------------------------------------------------------
' 1. BASE FINDING METRIC GROUP
' ---------------------------------------------------------------------

' Technical Impact (TI)
Function GetVal_TI(code As String) As Double
    Select Case UCase(Trim(code))
        Case "C": GetVal_TI = 1.0
        Case "H": GetVal_TI = 0.9
        Case "M": GetVal_TI = 0.6
        Case "L": GetVal_TI = 0.3
        Case "N": GetVal_TI = 0.0
        Case "D": GetVal_TI = 0.6
        Case "UK": GetVal_TI = 0.5
        Case "NA": GetVal_TI = 1.0
        Case Else: GetVal_TI = 0
    End Select
End Function

' Acquired Privilege (AP)
Function GetVal_AP(code As String) As Double
    Select Case UCase(Trim(code))
        Case "A": GetVal_AP = 1.0
        Case "P": GetVal_AP = 0.9
        Case "RU": GetVal_AP = 0.7
        Case "L": GetVal_AP = 0.6
        Case "N": GetVal_AP = 0.1
        Case "D": GetVal_AP = 0.7
        Case "UK": GetVal_AP = 0.5
        Case "NA": GetVal_AP = 1.0
        Case Else: GetVal_AP = 0
    End Select
End Function

' Acquired Privilege Layer (AL)
Function GetVal_AL(code As String) As Double
    Select Case UCase(Trim(code))
        Case "A": GetVal_AL = 1.0
        Case "S": GetVal_AL = 0.9
        Case "N": GetVal_AL = 0.7
        Case "E": GetVal_AL = 1.0
        Case "D": GetVal_AL = 0.9
        Case "UK": GetVal_AL = 0.5
        Case "NA": GetVal_AL = 1.0
        Case Else: GetVal_AL = 0
    End Select
End Function

' Internal Control Effectiveness (IC)
Function GetVal_IC(code As String) As Double
    Select Case UCase(Trim(code))
        Case "N": GetVal_IC = 1.0
        Case "L": GetVal_IC = 0.9
        Case "M": GetVal_IC = 0.7
        Case "I": GetVal_IC = 0.5
        Case "B": GetVal_IC = 0.3
        Case "C": GetVal_IC = 0.0
        Case "D": GetVal_IC = 0.6
        Case "UK": GetVal_IC = 0.5
        Case "NA": GetVal_IC = 1.0
        Case Else: GetVal_IC = 0
    End Select
End Function

' Finding Confidence (FC)
Function GetVal_FC(code As String) As Double
    Select Case UCase(Trim(code))
        Case "T": GetVal_FC = 1.0
        Case "LT": GetVal_FC = 0.8
        Case "F": GetVal_FC = 0.0
        Case "D": GetVal_FC = 0.8
        Case "UK": GetVal_FC = 0.5
        Case "NA": GetVal_FC = 1.0
        Case Else: GetVal_FC = 0
    End Select
End Function


' ---------------------------------------------------------------------
' 2. ATTACK SURFACE METRIC GROUP
' ---------------------------------------------------------------------

' Required Privilege (RP)
Function GetVal_RP(code As String) As Double
    Select Case UCase(Trim(code))
        Case "N": GetVal_RP = 1.0
        Case "L": GetVal_RP = 0.9
        Case "RU": GetVal_RP = 0.7
        Case "P": GetVal_RP = 0.6
        Case "A": GetVal_RP = 0.1
        Case "D": GetVal_RP = 0.7
        Case "UK": GetVal_RP = 0.5
        Case "NA": GetVal_RP = 1.0
        Case Else: GetVal_RP = 0
    End Select
End Function

' Required Privilege Layer (RL)
Function GetVal_RL(code As String) As Double
    Select Case UCase(Trim(code))
        Case "A": GetVal_RL = 1.0
        Case "S": GetVal_RL = 0.9
        Case "N": GetVal_RL = 0.7
        Case "E": GetVal_RL = 1.0
        Case "D": GetVal_RL = 0.9
        Case "UK": GetVal_RL = 0.5
        Case "NA": GetVal_RL = 1.0
        Case Else: GetVal_RL = 0
    End Select
End Function

' Access Vector (AV)
Function GetVal_AV(code As String) As Double
    Select Case UCase(Trim(code))
        Case "I": GetVal_AV = 1.0
        Case "R": GetVal_AV = 0.8   ' Intranet
        Case "V": GetVal_AV = 0.8    ' Private Network
        Case "A": GetVal_AV = 0.7   ' Adjacent Network
        Case "L": GetVal_AV = 0.5    ' Local
        Case "P": GetVal_AV = 0.2   ' Physical
        Case "D": GetVal_AV = 0.75
        Case "U": GetVal_AV = 0.5
        Case "NA": GetVal_AV = 1.0
        Case Else: GetVal_AV = 0
    End Select
End Function

' Authentication Strength (AS)
Function GetVal_AS(code As String) As Double
    Select Case UCase(Trim(code))
        Case "S": GetVal_AS = 0.7
        Case "M": GetVal_AS = 0.8
        Case "W": GetVal_AS = 0.9
        Case "N": GetVal_AS = 1.0
        Case "D": GetVal_AS = 0.85
        Case "UK": GetVal_AS = 0.5
        Case "NA": GetVal_AS = 1.0
        Case Else: GetVal_AS = 0
    End Select
End Function

' Level of Interaction (IN)
Function GetVal_IN(code As String) As Double
    Select Case UCase(Trim(code))
        Case "A": GetVal_IN = 1.0    ' Automated
        Case "T": GetVal_IN = 0.9   ' Typical/Limited
        Case "M": GetVal_IN = 0.8    ' Moderate
        Case "O": GetVal_IN = 0.3    ' Opportunistic
        Case "H": GetVal_IN = 0.1    ' High
        Case "NI": GetVal_IN = 0.0    ' No interaction
        Case "D": GetVal_IN = 0.55
        Case "UK": GetVal_IN = 0.5
        Case "NA": GetVal_IN = 1.0
        Case Else: GetVal_IN = 0
    End Select
End Function

' Deployment Scope (SC)
Function GetVal_SC(code As String) As Double
    Select Case UCase(Trim(code))
        Case "A": GetVal_SC = 1.0    ' All
        Case "M": GetVal_SC = 0.9    ' Moderate
        Case "R": GetVal_SC = 0.5    ' Rare
        Case "P": GetVal_SC = 0.1   ' Potentially Reachable
        Case "D": GetVal_SC = 0.7
        Case "UK": GetVal_SC = 0.5
        Case "NA": GetVal_SC = 1.0
        Case Else: GetVal_SC = 0
    End Select
End Function


' ---------------------------------------------------------------------
' 3. ENVIRONMENTAL METRIC GROUP
' ---------------------------------------------------------------------

' Business Impact (BI)
Function GetVal_BI(code As String) As Double
    Select Case UCase(Trim(code))
        Case "C": GetVal_BI = 1.0
        Case "H": GetVal_BI = 0.9
        Case "M": GetVal_BI = 0.6
        Case "L": GetVal_BI = 0.3
        Case "N": GetVal_BI = 0.0
        Case "D": GetVal_BI = 0.6
        Case "UK": GetVal_BI = 0.5
        Case "NA": GetVal_BI = 1.0
        Case Else: GetVal_BI = 0
    End Select
End Function

' Likelihood of Discovery (DI)
Function GetVal_DI(code As String) As Double
    Select Case UCase(Trim(code))
        Case "H": GetVal_DI = 1.0
        Case "M": GetVal_DI = 0.6
        Case "L": GetVal_DI = 0.2
        Case "D": GetVal_DI = 0.6
        Case "UK": GetVal_DI = 0.5
        Case "NA": GetVal_DI = 1.0
        Case Else: GetVal_DI = 0
    End Select
End Function

' Likelihood of Exploit (EX)
Function GetVal_EX(code As String) As Double
    Select Case UCase(Trim(code))
        Case "H": GetVal_EX = 1.0
        Case "M": GetVal_EX = 0.6
        Case "L": GetVal_EX = 0.2
        Case "N": GetVal_EX = 0.0
        Case "D": GetVal_EX = 0.6
        Case "UK": GetVal_EX = 0.5
        Case "NA": GetVal_EX = 1.0
        Case Else: GetVal_EX = 0
    End Select
End Function

' External Control Effectiveness (EC)
Function GetVal_EC(code As String) As Double
    Select Case UCase(Trim(code))
        Case "N": GetVal_EC = 1.0
        Case "L": GetVal_EC = 0.9
        Case "M": GetVal_EC = 0.7
        Case "I": GetVal_EC = 0.5
        Case "B": GetVal_EC = 0.3
        Case "C": GetVal_EC = 0.1
        Case "D": GetVal_EC = 0.6
        Case "UK": GetVal_EC = 0.5
        Case "NA": GetVal_EC = 1.0
        Case Else: GetVal_EC = 0
    End Select
End Function

' Prevalence (P)
Function GetVal_P(code As String) As Double
    Select Case UCase(Trim(code))
        Case "W": GetVal_P = 1.0    ' Widespread
        Case "H": GetVal_P = 0.9    ' High
        Case "C": GetVal_P = 0.8    ' Common
        Case "L": GetVal_P = 0.7    ' Limited
        Case "D": GetVal_P = 0.85
        Case "UK": GetVal_P = 0.5
        Case "NA": GetVal_P = 1.0
        Case Else: GetVal_P = 0
    End Select
End Function