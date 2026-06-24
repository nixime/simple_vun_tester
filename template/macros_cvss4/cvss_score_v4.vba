Option VBASupport 1
Option Explicit

' ==============================================================================
' GLOBAL VARIABLES
' ==============================================================================
' This public variable acts as a high-speed cached flag for your parsing engine
Public IsExcelEnvironment As Boolean
Public IsEnvInitialized   As Boolean

' Define the fixed positions for each metric
Private Enum CvsIndex
    idx_AV = 0
    idx_AC = 1
    idx_AT = 2
    idx_PR = 3
    idx_UI = 4
    idx_VC = 5
    idx_VI = 6
    idx_VA = 7
    idx_SC = 8
    idx_SI = 9
    idx_SA = 10
    idx_E  = 11
    idx_CR = 12
    idx_IR = 13
    idx_AR = 14
End Enum

' ==============================================================================
' CORE UTILITIES
' ==============================================================================
Public Function InitializeEnvironment() As Boolean 
    If Not IsEnvInitialized Then
        On Error Resume Next
        Dim testApp As Object
        Set testApp = Application
        
        If Err.Number = 0 And Not testApp Is Nothing Then
            InitializeEnvironment = True
        Else
            InitializeEnvironment = False
        End If
        
        On Error GoTo 0
        IsEnvInitialized = True
    End If
End Function

Private Function CreateDictionary() As Object
    On Error Resume Next
    If IsExcelEnvironment Then
        Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Else
        GlobalScope.BasicLibraries.loadLibrary("ScriptForge")
        Set CreateDictionary = CreateScriptService("Dictionary")
    End If
    On Error GoTo 0
End Function

Private Function UpdateDictValue(ByRef dict As Object, key As String, val As String)
    If IsExcelEnvironment Then
        dict(key) = val
    Else
        dict.ReplaceItem(key, val)
    End If
End Function

Public Function Round1(Value As Double) As Double
    ' Add a tiny epsilon to correct floating-point underflow/overflow precision issues
    ' and ensure standard 1-decimal rounding matches the CVSS specification
    Round1 = Int((Value + 0.0000001) * 10 + 0.5) / 10
End Function

Function GetMax(val1 As Double, val2 As Double) As Double
    If val1 > val2 Then
        GetMax = val1
    Else
        GetMax = val2
    End If
End Function

' ==============================================================================
' STATIC JSON DATA FUNCTIONS
' ==============================================================================
Function GetLevelValue(category As String, key As String) As Double
    ' Default return value if category or key is not found
    GetLevelValue = 0#
    
    Select Case UCase(category)
        Case "AV": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "A": GetLevelValue = 0.1
            Case "L": GetLevelValue = 0.2: Case "P": GetLevelValue = 0.3
        End Select
         
        Case "PR": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "L": GetLevelValue = 0.1
            Case "H": GetLevelValue = 0.2
        End Select
        
        Case "UI": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "P": GetLevelValue = 0.1
            Case "A": GetLevelValue = 0.2
        End Select
        
        Case "AC": Select Case UCase(key)
            Case "L": GetLevelValue = 0.0: Case "H": GetLevelValue = 0.1
        End Select
        
        Case "AT": Select Case UCase(key)
            Case "N": GetLevelValue = 0.0: Case "P": GetLevelValue = 0.1
        End Select
        
        Case "VC", "VI", "VA"
            Select Case UCase(key)
                Case "H": GetLevelValue = 0.0: Case "L": GetLevelValue = 0.1
                Case "N": GetLevelValue = 0.2
            End Select
            
        Case "SC", "SI", "SA"
            Select Case UCase(key)
                Case "S": GetLevelValue = 0.0: Case "H": GetLevelValue = 0.1
                Case "L": GetLevelValue = 0.2: Case "N": GetLevelValue = 0.3
            End Select
            
        Case "CR", "IR", "AR"
            Select Case UCase(key)
                Case "H": GetLevelValue = 0.0: Case "M": GetLevelValue = 0.1
                Case "L": GetLevelValue = 0.2
            End Select
            
        Case "E": Select Case UCase(key)
            Case "U": GetLevelValue = 0.2: Case "P": GetLevelValue = 0.1
            Case "A": GetLevelValue = 0.0
        End Select
    End Select
End Function

Function GetMaxSeverity(category As String, key As Integer, Optional subkey As Integer) As Integer
    Select Case UCase(category)
        Case "EQ1"
            Select Case key
                Case 0: GetMaxSeverity = 1
                Case 1: GetMaxSeverity = 4
                Case 2: GetMaxSeverity = 5
            End Select
            
        Case "EQ2"
            Select Case key
                Case 0: GetMaxSeverity = 1
                Case 1: GetMaxSeverity = 2
            End Select

        Case "EQ3EQ6"
            Select Case key
                Case 0
                    Select Case subkey
                        Case 0: GetMaxSeverity = 7
                        Case 1: GetMaxSeverity = 6
                    End Select
                Case 1: GetMaxSeverity = 8
                Case 2: GetMaxSeverity = 10
            End Select

        Case "EQ4"
            Select Case key
                Case 0: GetMaxSeverity = 6
                Case 1: GetMaxSeverity = 5
                Case 2: GetMaxSeverity = 4
            End Select
    End Select
End Function

Function GetMacroBaseValue(macroVector As String) As Double
    GetMacroBaseValue = -1
    Select Case macroVector
        Case "000000": GetMacroBaseValue = 10
        Case "000001": GetMacroBaseValue = 9.9
        Case "000010": GetMacroBaseValue = 9.8
        Case "000011": GetMacroBaseValue = 9.5
        Case "000020": GetMacroBaseValue = 9.5
        Case "000021": GetMacroBaseValue = 9.2
        Case "000100": GetMacroBaseValue = 10
        Case "000101": GetMacroBaseValue = 9.6
        Case "000110": GetMacroBaseValue = 9.3
        Case "000111": GetMacroBaseValue = 8.7
        Case "000120": GetMacroBaseValue = 9.1
        Case "000121": GetMacroBaseValue = 8.1
        Case "000200": GetMacroBaseValue = 9.3
        Case "000201": GetMacroBaseValue = 9
        Case "000210": GetMacroBaseValue = 8.9
        Case "000211": GetMacroBaseValue = 8
        Case "000220": GetMacroBaseValue = 8.1
        Case "000221": GetMacroBaseValue = 6.8
        Case "001000": GetMacroBaseValue = 9.8
        Case "001001": GetMacroBaseValue = 9.5
        Case "001010": GetMacroBaseValue = 9.5
        Case "001011": GetMacroBaseValue = 9.2
        Case "001020": GetMacroBaseValue = 9
        Case "001021": GetMacroBaseValue = 8.4
        Case "001100": GetMacroBaseValue = 9.3
        Case "001101": GetMacroBaseValue = 9.2
        Case "001110": GetMacroBaseValue = 8.9
        Case "001111": GetMacroBaseValue = 8.1
        Case "001120": GetMacroBaseValue = 8.1
        Case "001121": GetMacroBaseValue = 6.5
        Case "001200": GetMacroBaseValue = 8.8
        Case "001201": GetMacroBaseValue = 8
        Case "001210": GetMacroBaseValue = 7.8
        Case "001211": GetMacroBaseValue = 7
        Case "001220": GetMacroBaseValue = 6.9
        Case "001221": GetMacroBaseValue = 4.8
        Case "002001": GetMacroBaseValue = 9.2
        Case "002011": GetMacroBaseValue = 8.2
        Case "002021": GetMacroBaseValue = 7.2
        Case "002101": GetMacroBaseValue = 7.9
        Case "002111": GetMacroBaseValue = 6.9
        Case "002121": GetMacroBaseValue = 5
        Case "002201": GetMacroBaseValue = 6.9
        Case "002211": GetMacroBaseValue = 5.5
        Case "002221": GetMacroBaseValue = 2.7
        Case "010000": GetMacroBaseValue = 9.9
        Case "010001": GetMacroBaseValue = 9.7
        Case "010010": GetMacroBaseValue = 9.5
        Case "010011": GetMacroBaseValue = 9.2
        Case "010020": GetMacroBaseValue = 9.2
        Case "010021": GetMacroBaseValue = 8.5
        Case "010100": GetMacroBaseValue = 9.5
        Case "010101": GetMacroBaseValue = 9.1
        Case "010110": GetMacroBaseValue = 9
        Case "010111": GetMacroBaseValue = 8.3
        Case "010120": GetMacroBaseValue = 8.4
        Case "010121": GetMacroBaseValue = 7.1
        Case "010200": GetMacroBaseValue = 9.2
        Case "010201": GetMacroBaseValue = 8.1
        Case "010210": GetMacroBaseValue = 8.2
        Case "010211": GetMacroBaseValue = 7.1
        Case "010220": GetMacroBaseValue = 7.2
        Case "010221": GetMacroBaseValue = 5.3
        Case "011000": GetMacroBaseValue = 9.5
        Case "011001": GetMacroBaseValue = 9.3
        Case "011010": GetMacroBaseValue = 9.2
        Case "011011": GetMacroBaseValue = 8.5
        Case "011020": GetMacroBaseValue = 8.5
        Case "011021": GetMacroBaseValue = 7.3
        Case "011100": GetMacroBaseValue = 9.2
        Case "011101": GetMacroBaseValue = 8.2
        Case "011110": GetMacroBaseValue = 8
        Case "011111": GetMacroBaseValue = 7.2
        Case "011120": GetMacroBaseValue = 7
        Case "011121": GetMacroBaseValue = 5.9
        Case "011200": GetMacroBaseValue = 8.4
        Case "011201": GetMacroBaseValue = 7
        Case "011210": GetMacroBaseValue = 7.1
        Case "011211": GetMacroBaseValue = 5.2
        Case "011220": GetMacroBaseValue = 5
        Case "011221": GetMacroBaseValue = 3
        Case "012001": GetMacroBaseValue = 8.6
        Case "012011": GetMacroBaseValue = 7.5
        Case "012021": GetMacroBaseValue = 5.2
        Case "012101": GetMacroBaseValue = 7.1
        Case "012111": GetMacroBaseValue = 5.2
        Case "012121": GetMacroBaseValue = 2.9
        Case "012201": GetMacroBaseValue = 6.3
        Case "012211": GetMacroBaseValue = 2.9
        Case "012221": GetMacroBaseValue = 1.7
        Case "100000": GetMacroBaseValue = 9.8
        Case "100001": GetMacroBaseValue = 9.5
        Case "100010": GetMacroBaseValue = 9.4
        Case "100011": GetMacroBaseValue = 8.7
        Case "100020": GetMacroBaseValue = 9.1
        Case "100021": GetMacroBaseValue = 8.1
        Case "100100": GetMacroBaseValue = 9.4
        Case "100101": GetMacroBaseValue = 8.9
        Case "100110": GetMacroBaseValue = 8.6
        Case "100111": GetMacroBaseValue = 7.4
        Case "100120": GetMacroBaseValue = 7.7
        Case "100121": GetMacroBaseValue = 6.4
        Case "100200": GetMacroBaseValue = 8.7
        Case "100201": GetMacroBaseValue = 7.5
        Case "100210": GetMacroBaseValue = 7.4
        Case "100211": GetMacroBaseValue = 6.3
        Case "100220": GetMacroBaseValue = 6.3
        Case "100221": GetMacroBaseValue = 4.9
        Case "101000": GetMacroBaseValue = 9.4
        Case "101001": GetMacroBaseValue = 8.9
        Case "101010": GetMacroBaseValue = 8.8
        Case "101011": GetMacroBaseValue = 7.7
        Case "101020": GetMacroBaseValue = 7.6
        Case "101021": GetMacroBaseValue = 6.7
        Case "101100": GetMacroBaseValue = 8.6
        Case "101101": GetMacroBaseValue = 7.6
        Case "101110": GetMacroBaseValue = 7.4
        Case "101111": GetMacroBaseValue = 5.8
        Case "101120": GetMacroBaseValue = 5.9
        Case "101121": GetMacroBaseValue = 5
        Case "101200": GetMacroBaseValue = 7.2
        Case "101201": GetMacroBaseValue = 5.7
        Case "101210": GetMacroBaseValue = 5.7
        Case "101211": GetMacroBaseValue = 5.2
        Case "101220": GetMacroBaseValue = 5.2
        Case "101221": GetMacroBaseValue = 2.5
        Case "102001": GetMacroBaseValue = 8.3
        Case "102011": GetMacroBaseValue = 7
        Case "102021": GetMacroBaseValue = 5.4
        Case "102101": GetMacroBaseValue = 6.5
        Case "102111": GetMacroBaseValue = 5.8
        Case "102121": GetMacroBaseValue = 2.6
        Case "102201": GetMacroBaseValue = 5.3
        Case "102211": GetMacroBaseValue = 2.1
        Case "102221": GetMacroBaseValue = 1.3
        Case "110000": GetMacroBaseValue = 9.5
        Case "110001": GetMacroBaseValue = 9
        Case "110010": GetMacroBaseValue = 8.8
        Case "110011": GetMacroBaseValue = 7.6
        Case "110020": GetMacroBaseValue = 7.6
        Case "110021": GetMacroBaseValue = 7
        Case "110100": GetMacroBaseValue = 9
        Case "110101": GetMacroBaseValue = 7.7
        Case "110110": GetMacroBaseValue = 7.5
        Case "110111": GetMacroBaseValue = 6.2
        Case "110120": GetMacroBaseValue = 6.1
        Case "110121": GetMacroBaseValue = 5.3
        Case "110200": GetMacroBaseValue = 7.7
        Case "110201": GetMacroBaseValue = 6.6
        Case "110210": GetMacroBaseValue = 6.8
        Case "110211": GetMacroBaseValue = 5.9
        Case "110220": GetMacroBaseValue = 5.2
        Case "110221": GetMacroBaseValue = 3
        Case "111000": GetMacroBaseValue = 8.9
        Case "111001": GetMacroBaseValue = 7.8
        Case "111010": GetMacroBaseValue = 7.6
        Case "111011": GetMacroBaseValue = 6.7
        Case "111020": GetMacroBaseValue = 6.2
        Case "111021": GetMacroBaseValue = 5.8
        Case "111100": GetMacroBaseValue = 7.4
        Case "111101": GetMacroBaseValue = 5.9
        Case "111110": GetMacroBaseValue = 5.7
        Case "111111": GetMacroBaseValue = 5.7
        Case "111120": GetMacroBaseValue = 4.7
        Case "111121": GetMacroBaseValue = 2.3
        Case "111200": GetMacroBaseValue = 6.1
        Case "111201": GetMacroBaseValue = 5.2
        Case "111210": GetMacroBaseValue = 5.7
        Case "111211": GetMacroBaseValue = 2.9
        Case "111220": GetMacroBaseValue = 2.4
        Case "111221": GetMacroBaseValue = 1.6
        Case "112001": GetMacroBaseValue = 7.1
        Case "112011": GetMacroBaseValue = 5.9
        Case "112021": GetMacroBaseValue = 3
        Case "112101": GetMacroBaseValue = 5.8
        Case "112111": GetMacroBaseValue = 2.6
        Case "112121": GetMacroBaseValue = 1.5
        Case "112201": GetMacroBaseValue = 2.3
        Case "112211": GetMacroBaseValue = 1.3
        Case "112221": GetMacroBaseValue = 0.6
        Case "200000": GetMacroBaseValue = 9.3
        Case "200001": GetMacroBaseValue = 8.7
        Case "200010": GetMacroBaseValue = 8.6
        Case "200011": GetMacroBaseValue = 7.2
        Case "200020": GetMacroBaseValue = 7.5
        Case "200021": GetMacroBaseValue = 5.8
        Case "200100": GetMacroBaseValue = 8.6
        Case "200101": GetMacroBaseValue = 7.4
        Case "200110": GetMacroBaseValue = 7.4
        Case "200111": GetMacroBaseValue = 6.1
        Case "200120": GetMacroBaseValue = 5.6
        Case "200121": GetMacroBaseValue = 3.4
        Case "200200": GetMacroBaseValue = 7
        Case "200201": GetMacroBaseValue = 5.4
        Case "200210": GetMacroBaseValue = 5.2
        Case "200211": GetMacroBaseValue = 4
        Case "200220": GetMacroBaseValue = 4
        Case "200221": GetMacroBaseValue = 2.2
        Case "201000": GetMacroBaseValue = 8.5
        Case "201001": GetMacroBaseValue = 7.5
        Case "201010": GetMacroBaseValue = 7.4
        Case "201011": GetMacroBaseValue = 5.5
        Case "201020": GetMacroBaseValue = 6.2
        Case "201021": GetMacroBaseValue = 5.1
        Case "201100": GetMacroBaseValue = 7.2
        Case "201101": GetMacroBaseValue = 5.7
        Case "201110": GetMacroBaseValue = 5.5
        Case "201111": GetMacroBaseValue = 4.1
        Case "201120": GetMacroBaseValue = 4.6
        Case "201121": GetMacroBaseValue = 1.9
        Case "201200": GetMacroBaseValue = 5.3
        Case "201201": GetMacroBaseValue = 3.6
        Case "201210": GetMacroBaseValue = 3.4
        Case "201211": GetMacroBaseValue = 1.9
        Case "201220": GetMacroBaseValue = 1.9
        Case "201221": GetMacroBaseValue = 0.8
        Case "202001": GetMacroBaseValue = 6.4
        Case "202011": GetMacroBaseValue = 5.1
        Case "202021": GetMacroBaseValue = 2
        Case "202101": GetMacroBaseValue = 4.7
        Case "202111": GetMacroBaseValue = 2.1
        Case "202121": GetMacroBaseValue = 1.1
        Case "202201": GetMacroBaseValue = 2.4
        Case "202211": GetMacroBaseValue = 0.9
        Case "202221": GetMacroBaseValue = 0.4
        Case "210000": GetMacroBaseValue = 8.8
        Case "210001": GetMacroBaseValue = 7.5
        Case "210010": GetMacroBaseValue = 7.3
        Case "210011": GetMacroBaseValue = 5.3
        Case "210020": GetMacroBaseValue = 6
        Case "210021": GetMacroBaseValue = 5
        Case "210100": GetMacroBaseValue = 7.3
        Case "210101": GetMacroBaseValue = 5.5
        Case "210110": GetMacroBaseValue = 5.9
        Case "210111": GetMacroBaseValue = 4
        Case "210120": GetMacroBaseValue = 4.1
        Case "210121": GetMacroBaseValue = 2
        Case "210200": GetMacroBaseValue = 5.4
        Case "210201": GetMacroBaseValue = 4.3
        Case "210210": GetMacroBaseValue = 4.5
        Case "210211": GetMacroBaseValue = 2.2
        Case "210220": GetMacroBaseValue = 2
        Case "210221": GetMacroBaseValue = 1.1
        Case "211000": GetMacroBaseValue = 7.5
        Case "211001": GetMacroBaseValue = 5.5
        Case "211010": GetMacroBaseValue = 5.8
        Case "211011": GetMacroBaseValue = 4.5
        Case "211020": GetMacroBaseValue = 4
        Case "211021": GetMacroBaseValue = 2.1
        Case "211100": GetMacroBaseValue = 6.1
        Case "211101": GetMacroBaseValue = 5.1
        Case "211110": GetMacroBaseValue = 4.8
        Case "211111": GetMacroBaseValue = 1.8
        Case "211120": GetMacroBaseValue = 2
        Case "211121": GetMacroBaseValue = 0.9
        Case "211200": GetMacroBaseValue = 4.6
        Case "211201": GetMacroBaseValue = 1.8
        Case "211210": GetMacroBaseValue = 1.7
        Case "211211": GetMacroBaseValue = 0.7
        Case "211220": GetMacroBaseValue = 0.8
        Case "211221": GetMacroBaseValue = 0.2
        Case "212001": GetMacroBaseValue = 5.3
        Case "212011": GetMacroBaseValue = 2.4
        Case "212021": GetMacroBaseValue = 1.4
        Case "212101": GetMacroBaseValue = 2.4
        Case "212111": GetMacroBaseValue = 1.2
        Case "212121": GetMacroBaseValue = 0.5
        Case "212201": GetMacroBaseValue = 1
        Case "212211": GetMacroBaseValue = 0.3
        Case "212221": GetMacroBaseValue = 0.1
    End Select
End Function

' ==============================================================================
' LOOKUP & PARSING LOGIC
' ==============================================================================
Private Sub FillDefaultsArray(ByRef cvssData() As String)
    ' Pre-size the Array for the 15 metrics
    ReDim cvssData(0 To 14)
    
    ' Set fallback defaults matching your logic
    cvssData(idx_AV) = "N"
    cvssData(idx_AC) = "L"
    cvssData(idx_AT) = "N"
    cvssData(idx_PR) = "N"
    cvssData(idx_UI) = "N"
    cvssData(idx_VC) = "L"
    cvssData(idx_VI) = "L"
    cvssData(idx_VA) = "L"
    cvssData(idx_SC) = "N"
    cvssData(idx_SI) = "N"
    cvssData(idx_SA) = "N"
    cvssData(idx_E)  = "A" ' Default overrides applied immediately
    cvssData(idx_CR) = "H"
    cvssData(idx_IR) = "H"
    cvssData(idx_AR) = "H"
End Sub

Private Function ParseVectorToArray(ByVal VectorStr As String) As Variant
    Dim metrics(0 To 14) As String
    Dim parts() As String
    Dim pair() As String
    Dim i As Long
    
    metrics(idx_AV) = "N": metrics(idx_AC) = "L": metrics(idx_AT) = "N": metrics(idx_PR) = "N": metrics(idx_UI) = "N"
    metrics(idx_VC) = "L": metrics(idx_VI) = "L": metrics(idx_VA) = "L": metrics(idx_SC) = "N": metrics(idx_SI) = "N": metrics(idx_SA) = "N"
    metrics(idx_E)  = "X": metrics(idx_CR) = "X": metrics(idx_IR) = "X": metrics(idx_AR) = "X"
    
    VectorStr = UCase(Trim(VectorStr))
    If InStr(1, VectorStr, "INVALID") > 0 Or VectorStr = "" Then
        ParseVectorToArray = metrics
        Exit Function
    End If
    
    If Left(VectorStr, 9) = "CVSS:4.0/" Then VectorStr = Mid(VectorStr, 10)

    ' Walk raw vector and split elements directly into fields
    parts = Split(VectorStr, "/")
    For i = LBound(parts) To UBound(parts)
        pair = Split(parts(i), ":")
        If UBound(pair) = 1 Then
            Select Case pair(0)
                Case "AV": metrics(idx_AV) = pair(1)
                Case "AC": metrics(idx_AC) = pair(1)
                Case "AT": metrics(idx_AT) = pair(1)
                Case "PR": metrics(idx_PR) = pair(1)
                Case "UI": metrics(idx_UI) = pair(1)
                Case "VC": metrics(idx_VC) = pair(1)
                Case "VI": metrics(idx_VI) = pair(1)
                Case "VA": metrics(idx_VA) = pair(1)
                Case "SC": metrics(idx_SC) = pair(1)
                Case "SI": metrics(idx_SI) = pair(1)
                Case "SA": metrics(idx_SA) = pair(1)
                Case "E":  metrics(idx_E)  = pair(1)
                Case "CR": metrics(idx_CR) = pair(1)
                Case "IR": metrics(idx_IR) = pair(1)
                Case "AR": metrics(idx_AR) = pair(1)
            End Select
        End If
    Next i

    If metrics(idx_E) = "X"  Then metrics(idx_E) = "A"
    If metrics(idx_CR) = "X" Then metrics(idx_CR) = "H"
    If metrics(idx_IR) = "X" Then metrics(idx_IR) = "H"
    If metrics(idx_AR) = "X" Then metrics(idx_AR) = "H"
    
    ParseVectorToArray = metrics
End Function

' ==============================================================================
' CVSS 4.0 EQUATION LEVELS
' ==============================================================================
Private Function getEqLevel1(ByRef vecArray() As String) As Integer
    If vecArray(idx_AV) = "N" And vecArray(idx_PR) = "N" And vecArray(idx_UI) = "N" Then
        getEqLevel1 = 0
    ElseIf (vecArray(idx_AV) = "N" Or vecArray(idx_PR) = "N" Or vecArray(idx_UI) = "N") And Not vecArray(idx_AV) = "P" Then
        getEqLevel1 = 1
    Else
        getEqLevel1 = 2
    End If
End Function

Private Function getEqLevel2(ByRef vecArray() As String) As Integer
    If vecArray(idx_AC) = "L" And vecArray(idx_AT) = "N" Then
        getEqLevel2 = 0
    Else
        getEqLevel2 = 1
    End If
End Function

Private Function getEqLevel3(ByRef vecArray() As String) As Integer
    If vecArray(idx_VC) = "H" And vecArray(idx_VI) = "H" Then
        getEqLevel3 = 0
    ElseIf vecArray(idx_VC) = "H" Or vecArray(idx_VI) = "H" Or vecArray(idx_VA) = "H" Then
        getEqLevel3 = 1
    Else
        getEqLevel3 = 2
    End If
End Function

Private Function getEqLevel4(ByRef vecArray() As String) As Integer
    If vecArray(idx_SI) = "S" Or vecArray(idx_SA) = "S" Then
        getEqLevel4 = 0
    ElseIf vecArray(idx_SC) = "H" Or vecArray(idx_SI) = "H" Or vecArray(idx_SA) = "H" Then
        getEqLevel4 = 1
    Else
        getEqLevel4 = 2
    End If
End Function

Private Function getEqLevel5(ByRef vecArray() As String) As Integer
    If vecArray(idx_E) = "A" Or vecArray(idx_E) = "X" Then
        getEqLevel5 = 0
    ElseIf vecArray(idx_E) = "P" Then
        getEqLevel5 = 1
    Else
        getEqLevel5 = 2
    End If
End Function

Private Function getEqLevel6(ByRef vecArray() As String) As Integer
    Dim cr As String, vc As String, ir As String, vi As String, ar As String, va As String
    cr = vecArray(idx_CR): vc = vecArray(idx_VC): ir = vecArray(idx_IR): vi = vecArray(idx_VI): ar = vecArray(idx_AR): va = vecArray(idx_VA)

    If cr = "X" Then cr = "H"
    If ir = "X" Then ir = "H"
    If ar = "X" Then ar = "H"
    
    If (cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H") Then
        getEqLevel6 = 0
    Else
        getEqLevel6 = 1
    End If
End Function

Private Function getJointEqLevel36(ByRef vecArray() As String) As String
    Dim cr As String, vc As String, ir As String, vi As String, ar As String, va As String
    cr = vecArray(idx_CR): vc = vecArray(idx_VC): ir = vecArray(idx_IR): vi = vecArray(idx_VI): ar = vecArray(idx_AR): va = vecArray(idx_VA)
    
    If vc = "H" And vi = "H" And (cr = "H" Or ir = "H" Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "00"
    ElseIf vc = "H" And vi = "H" And Not (cr = "H" Or ir = "H") And Not (ar = "H" And va = "H") Then
        getJointEqLevel36 = "01"
    ElseIf Not (vc = "H" And vi = "H") And (vc = "H" Or vi = "H" Or va = "H") And _
           ((cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "10"
    ElseIf (vc = "H" And vi = "H") And (vc = "H" Or vi = "H" Or va = "H") And _
           (Not (cr = "H" And vc = "H") And Not (ir = "H" And vi = "H") And Not (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "11"
    ElseIf Not (vc = "H" Or vi = "H" Or va = "H") And _
           ((cr = "H" And vc = "H") Or (ir = "H" And vi = "H") Or (ar = "H" And va = "H")) Then
        getJointEqLevel36 = "20"
    Else
        getJointEqLevel36 = "21"
    End If
End Function

Private Function get_severity_distance_vecArray(ByRef vecArray() As String, ByRef max_vecArray() As String, ByVal key As String) As Double
    Dim idx As Integer
    Select Case key
        Case "AV": idx = idx_AV: Case "AC": idx = idx_AC: Case "AT": idx = idx_AT
        Case "PR": idx = idx_PR: Case "UI": idx = idx_UI: Case "VC": idx = idx_VC
        Case "VI": idx = idx_VI: Case "VA": idx = idx_VA: Case "SC": idx = idx_SC
        Case "SI": idx = idx_SI: Case "SA": idx = idx_SA: Case "E":  idx = idx_E
        Case "CR": idx = idx_CR: Case "IR": idx = idx_IR: Case "AR": idx = idx_AR
    End Select
    get_severity_distance_vecArray = GetLevelValue(key, vecArray(idx)) - GetLevelValue(key, max_vecArray(idx))
End Function

Function getEQMaxes(lookup As String, eqIndex As Integer, Optional subEqIndex As Integer) As Variant
    Dim lookupKey As String
    Dim lookupSubKey As String
    lookupKey = Mid(lookup, eqIndex, 1)
    getEQMaxes = Array()
    Select Case eqIndex
        Case 1
            Select Case lookupKey
                Case "0": getEQMaxes = Array("AV:N/PR:N/UI:N/")
                Case "1": getEQMaxes = Array("AV:A/PR:N/UI:N/", "AV:N/PR:L/UI:N/", "AV:N/PR:N/UI:P/")
                Case "2": getEQMaxes = Array("AV:P/PR:N/UI:N/", "AV:A/PR:L/UI:P/")
            End Select
        Case 2
            Select Case lookupKey
                Case "0": getEQMaxes = Array("AC:L/AT:N/")
                Case "1": getEQMaxes = Array("AC:H/AT:N/", "AC:L/AT:P/")
            End Select
        Case 3
            lookupSubKey = Mid(lookup, subEqIndex, 1)
            Select Case lookupKey
                Case "0"
                    If lookupSubKey = "0" Then
                        getEQMaxes = Array("VC:H/VI:H/VA:H/CR:H/IR:H/AR:H/")
                    ElseIf lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:H/VI:H/VA:L/CR:M/IR:M/AR:H/", "VC:H/VI:H/VA:H/CR:M/IR:M/AR:M/")
                    End If
                Case "1"
                    If lookupSubKey = "0" Then
                        getEQMaxes = Array("VC:L/VI:H/VA:H/CR:H/IR:H/AR:H/", "VC:H/VI:L/VA:H/CR:H/IR:H/AR:H/")
                    ElseIf lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:L/VI:H/VA:L/CR:H/IR:M/AR:H/", "VC:L/VI:H/VA:H/CR:H/IR:M/AR:M/", "VC:H/VI:L/VA:H/CR:M/IR:H/AR:M/", "VC:H/VI:L/VA:L/CR:M/IR:H/AR:H/", "VC:L/VI:L/VA:H/CR:H/IR:H/AR:M/")
                    End If
                Case "2"
                    If lookupSubKey = "1" Then
                        getEQMaxes = Array("VC:L/VI:L/VA:L/CR:H/IR:H/AR:H/")
                    End If
            End Select
        Case 4
            Select Case lookupKey
                Case "0": getEQMaxes = Array("SC:H/SI:S/SA:S/")
                Case "1": getEQMaxes = Array("SC:H/SI:H/SA:H/")
                Case "2": getEQMaxes = Array("SC:L/SI:L/SA:L/")
            End Select
        Case 5
            Select Case lookupKey
                Case "0": getEQMaxes = Array("E:A/")
                Case "1": getEQMaxes = Array("E:P/")
                Case "2": getEQMaxes = Array("E:U/")
            End Select
    End Select
End Function

' ==============================================================================
' PUBLIC INTERFACE
' ==============================================================================
Public Function GetMacroVector(ByVal VectorStr As String) As String
    Dim vecArray As Variant
    vecArray = ParseVectorToArray(VectorStr)
    GetMacroVector = CStr(getEqLevel1(vecArray)) & CStr(getEqLevel2(vecArray)) & _
                     CStr(getEqLevel3(vecArray)) & CStr(getEqLevel4(vecArray)) & _
                     CStr(getEqLevel5(vecArray)) & CStr(getEqLevel6(vecArray))
End Function

Public Function cvss_score(ByVal VectorStr As String) As Double
    Dim vecArray As Variant
    Dim IsNotApplicable As Boolean
    Dim macroVectorScore As Double

    On Error GoTo ErrorHandler

    vecArray = ParseVectorToArray(VectorStr)

    IsNotApplicable = True
    If vecArray(idx_VC) <> "N" Then IsNotApplicable = False
    If vecArray(idx_VI) <> "N" Then IsNotApplicable = False
    If vecArray(idx_VA) <> "N" Then IsNotApplicable = False
    If vecArray(idx_SC) <> "N" Then IsNotApplicable = False
    If vecArray(idx_SI) <> "N" Then IsNotApplicable = False
    If vecArray(idx_SA) <> "N" Then IsNotApplicable = False

    If IsNotApplicable Then
        cvss_score = 0.0
        Exit Function
    End If

    Dim macroValueResult As String
    Dim eq1 As Integer, eq2 As Integer, eq3 As Integer, eq4 As Integer, eq5 As Integer, eq6 As Integer
 
    macroValueResult = GetMacroVector(VectorStr)
    macroVectorScore = GetMacroBaseValue(macroValueResult)
    If macroVectorScore = -1 Then
        cvss_score = -1
        Exit Function
    End If

    eq1 = CInt(Mid(macroValueResult, 1, 1))
    eq2 = CInt(Mid(macroValueResult, 2, 1))
    eq3 = CInt(Mid(macroValueResult, 3, 1))
    eq4 = CInt(Mid(macroValueResult, 4, 1))
    eq5 = CInt(Mid(macroValueResult, 5, 1))
    eq6 = CInt(Mid(macroValueResult, 6, 1))

    Dim eq1_next_lower_macro As String: eq1_next_lower_macro = ""
    Dim eq2_next_lower_macro As String: eq2_next_lower_macro = ""
    Dim eq4_next_lower_macro As String: eq4_next_lower_macro = ""
    Dim eq5_next_lower_macro As String: eq5_next_lower_macro = ""

    If eq1 < 2 Then eq1_next_lower_macro = CStr(eq1 + 1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    If eq2 < 2 Then eq2_next_lower_macro = CStr(eq1) & CStr(eq2 + 1) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    If eq4 < 2 Then eq4_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4 + 1) & CStr(eq5) & CStr(eq6)
    If eq5 < 2 Then eq5_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5 + 1) & CStr(eq6)
    
    Dim eq3eq6_next_lower_macro As String: eq3eq6_next_lower_macro = ""
    Dim eq3eq6_next_lower_macro_left As String: eq3eq6_next_lower_macro_left = ""
    Dim eq3eq6_next_lower_macro_right As String: eq3eq6_next_lower_macro_right = ""
    
    If eq3 = 1 And eq6 = 1 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    ElseIf eq3 = 0 And eq6 = 1 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    ElseIf eq3 = 1 And eq6 = 0 Then
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
    ElseIf eq3 = 0 And eq6 = 0 Then
        eq3eq6_next_lower_macro_left = CStr(eq1) & CStr(eq2) & CStr(eq3) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
        eq3eq6_next_lower_macro_right = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6)
    Else
        eq3eq6_next_lower_macro = CStr(eq1) & CStr(eq2) & CStr(eq3 + 1) & CStr(eq4) & CStr(eq5) & CStr(eq6 + 1)
    End If

    Dim score_eq1_next_lower_macro As Variant: score_eq1_next_lower_macro = -1
    Dim score_eq2_next_lower_macro As Variant: score_eq2_next_lower_macro = -1
    Dim score_eq3eq6_next_lower_macro As Variant: score_eq3eq6_next_lower_macro = -1
    Dim score_eq3eq6_next_lower_macro_left As Variant: score_eq3eq6_next_lower_macro_left = -1
    Dim score_eq3eq6_next_lower_macro_right As Variant: score_eq3eq6_next_lower_macro_right = -1
    Dim score_eq4_next_lower_macro As Variant: score_eq4_next_lower_macro = -1
    Dim score_eq5_next_lower_macro As Variant: score_eq5_next_lower_macro = -1
    
    If eq1_next_lower_macro <> "" Then score_eq1_next_lower_macro = GetMacroBaseValue(eq1_next_lower_macro)
    If eq2_next_lower_macro <> "" Then score_eq2_next_lower_macro = GetMacroBaseValue(eq2_next_lower_macro)
    If eq4_next_lower_macro <> "" Then score_eq4_next_lower_macro = GetMacroBaseValue(eq4_next_lower_macro)
    If eq5_next_lower_macro <> "" Then score_eq5_next_lower_macro = GetMacroBaseValue(eq5_next_lower_macro)

    If eq3 = 0 And eq6 = 0 Then
        score_eq3eq6_next_lower_macro_left = GetMacroBaseValue(eq3eq6_next_lower_macro_left)
        score_eq3eq6_next_lower_macro_right = GetMacroBaseValue(eq3eq6_next_lower_macro_right)
        
        If score_eq3eq6_next_lower_macro_left < 0 Then score_eq3eq6_next_lower_macro_left = macroVectorScore
        If score_eq3eq6_next_lower_macro_right < 0 Then score_eq3eq6_next_lower_macro_right = macroVectorScore

        If Val(score_eq3eq6_next_lower_macro_left) > Val(score_eq3eq6_next_lower_macro_right) Then
            score_eq3eq6_next_lower_macro = score_eq3eq6_next_lower_macro_left
            eq3eq6_next_lower_macro = eq3eq6_next_lower_macro_left
        Else
            score_eq3eq6_next_lower_macro = score_eq3eq6_next_lower_macro_right
            eq3eq6_next_lower_macro = eq3eq6_next_lower_macro_right
        End If
    Else
        score_eq3eq6_next_lower_macro = GetMacroBaseValue(eq3eq6_next_lower_macro)
        If score_eq3eq6_next_lower_macro < 0 Then eq3eq6_next_lower_macro = ""
    End If

    Dim eq1_maxes As Variant, eq2_maxes As Variant, eq3_eq6_maxes As Variant, eq4_maxes As Variant, eq5_maxes As Variant
    eq1_maxes = getEQMaxes(macroValueResult, 1)
    eq2_maxes = getEQMaxes(macroValueResult, 2)
    eq3_eq6_maxes = getEQMaxes(macroValueResult, 3, 6)
    eq4_maxes = getEQMaxes(macroValueResult, 4)
    eq5_maxes = getEQMaxes(macroValueResult, 5)

    Dim max_vectors() As String
    Dim max_vectors_size As Integer: max_vectors_size = 0
    Dim eq1_max As Variant, eq2_max As Variant, eq3_eq6_max As Variant, eq4_max As Variant, eq5max As Variant
    Dim push_val As String

    For Each eq1_max In eq1_maxes
        For Each eq2_max In eq2_maxes
            For Each eq3_eq6_max In eq3_eq6_maxes
                For Each eq4_max In eq4_maxes
                    For Each eq5max In eq5_maxes
                        ReDim Preserve max_vectors(0 To max_vectors_size)
                        push_val = eq1_max & eq2_max & eq3_eq6_max & eq4_max & eq5max
                        max_vectors(max_vectors_size) = push_val
                        max_vectors_size = max_vectors_size + 1
                    Next eq5max
                Next eq4_max
            Next eq3_eq6_max
        Next eq2_max
    Next eq1_max

    Dim severity_distance_AV As Double, severity_distance_PR As Double, severity_distance_UI As Double
    Dim severity_distance_AC As Double, severity_distance_AT As Double
    Dim severity_distance_VC As Double, severity_distance_VI As Double, severity_distance_VA As Double
    Dim severity_distance_SC As Double, severity_distance_SI As Double, severity_distance_SA As Double
    Dim severity_distance_CR As Double, severity_distance_IR As Double, severity_distance_AR As Double
    Dim max_vector As Variant
    Dim max_vector_vecArray() As String

    For Each max_vector In max_vectors
        max_vector_vecArray = ParseVectorToArray(max_vector)
        
        severity_distance_AV = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "AV")
        severity_distance_PR = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "PR")
        severity_distance_UI = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "UI")
        severity_distance_AC = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "AC")
        severity_distance_AT = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "AT")
        severity_distance_VC = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "VC")
        severity_distance_VI = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "VI")
        severity_distance_VA = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "VA")
        severity_distance_SC = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "SC")
        severity_distance_SI = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "SI")
        severity_distance_SA = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "SA")
        severity_distance_CR = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "CR")
        severity_distance_IR = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "IR")
        severity_distance_AR = get_severity_distance_vecArray(vecArray, max_vector_vecArray, "AR")

        Dim less_than_zero As Boolean: less_than_zero = False
        Dim severity As Variant
        For Each severity In Array(severity_distance_AV, severity_distance_PR, severity_distance_UI, _
                                   severity_distance_AC, severity_distance_AT, severity_distance_VC, _
                                   severity_distance_VI, severity_distance_VA, severity_distance_SC, _
                                   severity_distance_SI, severity_distance_SA, severity_distance_CR, _
                                   severity_distance_IR, severity_distance_AR)
            If severity < 0 Then
                less_than_zero = True
            End If
        Next severity

        If Not less_than_zero Then
            Exit For
        End If
    Next max_vector

    Dim current_severity_distance_eq1 As Double
    Dim current_severity_distance_eq2 As Double
    Dim current_severity_distance_eq3eq6 As Double
    Dim current_severity_distance_eq4 As Double
    Dim current_severity_distance_eq5 As Double
    
    current_severity_distance_eq1 = severity_distance_AV + severity_distance_PR + severity_distance_UI
    current_severity_distance_eq2 = severity_distance_AC + severity_distance_AT
    current_severity_distance_eq3eq6 = severity_distance_VC + severity_distance_VI + severity_distance_VA + severity_distance_CR + severity_distance_IR + severity_distance_AR
    current_severity_distance_eq4 = severity_distance_SC + severity_distance_SI + severity_distance_SA
    current_severity_distance_eq5 = 0

    Dim available_distance_eq1 As Double: available_distance_eq1 = macroVectorScore - Val(score_eq1_next_lower_macro)
    Dim available_distance_eq2 As Double: available_distance_eq2 = macroVectorScore - Val(score_eq2_next_lower_macro)
    Dim available_distance_eq3eq6 As Double: available_distance_eq3eq6 = macroVectorScore - Val(score_eq3eq6_next_lower_macro)
    Dim available_distance_eq4 As Double: available_distance_eq4 = macroVectorScore - Val(score_eq4_next_lower_macro)
    Dim available_distance_eq5 As Double: available_distance_eq5 = macroVectorScore - Val(score_eq5_next_lower_macro)
    
    If available_distance_eq1 < 0 Then available_distance_eq1 = 0
    If available_distance_eq2 < 0 Then available_distance_eq2 = 0
    If available_distance_eq3eq6 < 0 Then available_distance_eq3eq6 = 0
    If available_distance_eq4 < 0 Then available_distance_eq4 = 0
    If available_distance_eq5 < 0 Then available_distance_eq5 = 0

    Dim n_existing_lower As Double: n_existing_lower = 0
    Dim normalized_severity_eq1 As Double: normalized_severity_eq1 = 0
    Dim normalized_severity_eq2 As Double: normalized_severity_eq2 = 0
    Dim normalized_severity_eq3eq6 As Double: normalized_severity_eq3eq6 = 0
    Dim normalized_severity_eq4 As Double: normalized_severity_eq4 = 0
    Dim normalized_severity_eq5 As Double: normalized_severity_eq5 = 0

    Dim maxSeverity_eq1 As Double, maxSeverity_eq2 As Double, maxSeverity_eq3eq6 As Double, maxSeverity_eq4 As Double
    maxSeverity_eq1 = GetMaxSeverity("eq1", eq1) * 0.1
    maxSeverity_eq2 = GetMaxSeverity("eq2", eq2) * 0.1
    maxSeverity_eq3eq6 = GetMaxSeverity("eq3eq6", eq3, eq6) * 0.1
    maxSeverity_eq4 = GetMaxSeverity("eq4", eq4) * 0.1

    Dim percent_to_next_eq1_severity As Double
    Dim percent_to_next_eq2_severity As Double
    Dim percent_to_next_eq3eq6_severity As Double
    Dim percent_to_next_eq4_severity As Double
    Dim percent_to_next_eq5_severity As Double

    If eq1_next_lower_macro <> "" And score_eq1_next_lower_macro >= 0 Then
        available_distance_eq1 = macroVectorScore - Val(score_eq1_next_lower_macro)
        If available_distance_eq1 < 0 Then available_distance_eq1 = 0

        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq1_severity = (current_severity_distance_eq1) / maxSeverity_eq1
        normalized_severity_eq1 = available_distance_eq1 * percent_to_next_eq1_severity
    End If
    
    If eq2_next_lower_macro <> "" And score_eq2_next_lower_macro >= 0 Then
        available_distance_eq2 = macroVectorScore - Val(score_eq2_next_lower_macro)
        If available_distance_eq2 < 0 Then available_distance_eq2 = 0

        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq2_severity = (current_severity_distance_eq2) / maxSeverity_eq2
        normalized_severity_eq2 = available_distance_eq2 * percent_to_next_eq2_severity
    End If

    If eq3eq6_next_lower_macro <> "" And score_eq3eq6_next_lower_macro >= 0 Then
        available_distance_eq3eq6 = macroVectorScore - Val(score_eq3eq6_next_lower_macro)
        If available_distance_eq3eq6 < 0 Then available_distance_eq3eq6 = 0

        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq3eq6_severity = (current_severity_distance_eq3eq6) / maxSeverity_eq3eq6
        normalized_severity_eq3eq6 = available_distance_eq3eq6 * percent_to_next_eq3eq6_severity
    End If

    If eq4_next_lower_macro <> "" And score_eq4_next_lower_macro >= 0 Then
        available_distance_eq4 = macroVectorScore - Val(score_eq4_next_lower_macro)
        If available_distance_eq4 < 0 Then available_distance_eq4 = 0

        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq4_severity = (current_severity_distance_eq4) / maxSeverity_eq4
        normalized_severity_eq4 = available_distance_eq4 * percent_to_next_eq4_severity
    End If
    
    If eq5_next_lower_macro <> "" And score_eq5_next_lower_macro >= 0 Then
        available_distance_eq5 = macroVectorScore - Val(score_eq5_next_lower_macro)
        If available_distance_eq5 < 0 Then available_distance_eq5 = 0

        n_existing_lower = n_existing_lower + 1
        percent_to_next_eq5_severity = 0
        normalized_severity_eq5 = available_distance_eq5 * percent_to_next_eq5_severity
    End If

    Dim mean_distance As Double
    If n_existing_lower = 0 Then
        mean_distance = 0
    Else
        mean_distance = (normalized_severity_eq1 + normalized_severity_eq2 + normalized_severity_eq3eq6 + normalized_severity_eq4 + normalized_severity_eq5) / n_existing_lower
    End If

    macroVectorScore = macroVectorScore - mean_distance
    If macroVectorScore < 0 Then macroVectorScore = 0.0
    If macroVectorScore > 10 Then macroVectorScore = 10.0
    
    cvss_score = Round1(macroVectorScore)
    Exit Function

ErrorHandler:
    cvss_score = -10
End Function
